import os
import json
import winshell
import sys
from PIL import Image, UnidentifiedImageError
import colorgram
import argparse
import traceback
import logging
import win32gui # type: ignore
import win32api # type: ignore
import win32con # type: ignore
from PIL import ImageWin # type: ignore # Used for Dib structure potentially, but GetDIBits is more direct
from ctypes import windll, byref, sizeof, c_ubyte, Structure, c_long, c_ushort, c_uint, c_int # Import c_int
import ctypes
from io import BytesIO

# Use constants from win32con
DI_NORMAL = win32con.DI_NORMAL
SM_CXICON = win32con.SM_CXICON
SM_CYICON = win32con.SM_CYICON

# --- Configuration and Paths (Keep as is) ---
# Color of the book cover in the template image (adjust if needed)
TEMPLATE_COVER_COLOR_RGB = (106, 156, 66) # Example: The specific green
COLOR_MATCH_TOLERANCE = 30 # How close a pixel color needs to be to be replaced

# Paths
SCRIPT_DIR = os.path.dirname(os.path.realpath(__file__))
TEMPLATE_PATH = os.path.join(SCRIPT_DIR, "book_template.png")
DESKTOP_PATH = winshell.desktop()
BACKUP_DIR = os.path.join(SCRIPT_DIR, "icon_backups")
BACKUP_FILE = os.path.join(BACKUP_DIR, "icon_backup.json")
GENERATED_ICONS_DIR = os.path.join(SCRIPT_DIR, "generated_icons")

# Desired output icon sizes (Windows uses multiple)
ICO_SIZES = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]

# Logging setup
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
# To see DEBUG messages for handles etc:
# logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')


# --- Helper Functions (ensure_dir_exists, rgb_color_distance, get_desktop_shortcuts remain the same) ---
def ensure_dir_exists(dir_path):
    os.makedirs(dir_path, exist_ok=True)

def rgb_color_distance(color1, color2):
    return sum([(c1 - c2) ** 2 for c1, c2 in zip(color1, color2)]) ** 0.5

def get_desktop_shortcuts():
    shortcuts = []
    try:
        for item in os.listdir(DESKTOP_PATH):
            if item.lower().endswith(".lnk"):
                shortcuts.append(os.path.join(DESKTOP_PATH, item))
    except FileNotFoundError:
        logging.error(f"Desktop path not found: {DESKTOP_PATH}")
    except Exception as e:
        logging.error(f"Error listing desktop items: {e}")
    return shortcuts

# --- get_original_icon_info (Keep the refined version from the previous response) ---
def get_original_icon_info(shortcut_path):
    """Gets the original icon location (path, index) from a shortcut."""
    try:
        with winshell.shortcut(shortcut_path) as lnk:
            icon_location = lnk.icon_location # This is often (path, index)
            path = None
            index = 0

            if isinstance(icon_location, (list, tuple)) and len(icon_location) >= 1:
                path = icon_location[0]
                if len(icon_location) > 1:
                    index = icon_location[1]
            elif isinstance(icon_location, str) and icon_location:
                # Handle "path,index" format
                parts = icon_location.split(',')
                path = parts[0].strip()
                index = int(parts[1].strip()) if len(parts) > 1 else 0
            # If path is still None or empty after checking icon_location, try the target
            if not path:
                target_path = lnk.path
                # Check if target exists and seems like a file that could have an icon
                if target_path and os.path.exists(target_path) and (target_path.lower().endswith((".exe", ".dll", ".ico")) or os.path.isfile(target_path)):
                     logging.info(f"No explicit icon set for {os.path.basename(shortcut_path)}, using target: {target_path}")
                     path = target_path
                     index = 0 # Usually index 0 for target's own icon
                else:
                    logging.warning(f"Cannot determine icon for {os.path.basename(shortcut_path)} - icon_location empty/invalid and target unusable ('{target_path}').")
                    return None, None

            # Resolve environment variables and check existence
            resolved_path = os.path.expandvars(path)
            if not os.path.exists(resolved_path):
                 logging.warning(f"Icon path '{resolved_path}' (from '{path}') does not exist for shortcut {os.path.basename(shortcut_path)}.")
                 return None, None

            return resolved_path, index
    except Exception as e:
        logging.error(f"Error reading shortcut '{os.path.basename(shortcut_path)}': {e}")
        # traceback.print_exc() # Uncomment for detailed debugging
        return None, None

# --- Corrected extract_icon_image ---
def extract_icon_image(icon_path, index):
    """
    Extracts the icon using win32gui for exe/dll, falling back slightly if needed.
    Corrected GetDIBits call.
    """
    if not icon_path or not os.path.exists(icon_path):
        logging.warning(f"Icon path does not exist: {icon_path}")
        return None

    try:
        if icon_path.lower().endswith(".ico"):
            img = Image.open(icon_path)
            try:
                # Attempt to load the largest frame if available
                if hasattr(img, 'info') and 'sizes' in img.info:
                     largest_size = max(img.info['sizes'])
                     if largest_size != (0,0):
                         img.size = largest_size
                         img.load() # Reload with the largest size selected
                else:
                    img.load() # Load default if sizes info not present
            except Exception as e:
                logging.warning(f"Could not determine/load largest size for ICO {os.path.basename(icon_path)}, using default: {e}")
                img.load() # Load default size if selection fails

            if img.mode != 'RGBA':
                img = img.convert('RGBA')
            return img

        elif icon_path.lower().endswith((".exe", ".dll")):
            logging.info(f"Extracting icon from {os.path.basename(icon_path)} at index {index} using win32gui.")

            try:
                large_icons, small_icons = win32gui.ExtractIconEx(icon_path, index, 1)
            except win32gui.error as e:
                 logging.error(f"win32gui.ExtractIconEx failed for '{os.path.basename(icon_path)}', index {index}. Error: {e}")
                 return None

            hicon = None
            icons_to_destroy = []

            if large_icons:
                hicon = large_icons[0]
                icons_to_destroy.extend(large_icons[1:])
                icons_to_destroy.extend(small_icons)
            elif small_icons:
                hicon = small_icons[0]
                logging.warning(f"Using small icon for {os.path.basename(icon_path)} as large icon was not available.")
                icons_to_destroy.extend(small_icons[1:])
            else:
                logging.error(f"No icon handles returned by ExtractIconEx for '{os.path.basename(icon_path)}', index {index}.")
                return None

            # Get system default icon size
            icon_x = win32api.GetSystemMetrics(SM_CXICON)
            icon_y = win32api.GetSystemMetrics(SM_CYICON)
            target_size = (icon_x, icon_y)

            hdc = None
            hdc_mem = None
            hbmp = None
            image = None # Initialize image to None

            try:
                hdc = win32gui.GetDC(0)
                hdc_mem = win32gui.CreateCompatibleDC(hdc)
                hbmp = win32gui.CreateCompatibleBitmap(hdc, target_size[0], target_size[1])

                if not hbmp:
                    logging.error(f"CreateCompatibleBitmap failed for {os.path.basename(icon_path)}.")
                    return None # Exit early, finally block will clean up handles

                # Select the bitmap into the memory DC
                # Store the old object to select back later (good practice)
                old_bmp = win32gui.SelectObject(hdc_mem, hbmp)

                # Draw the icon onto the bitmap context
                # Fill background with transparent or a known color first? Maybe not needed if icon draws fully.
                # win32gui.FillRect(hdc_mem, (0, 0, target_size[0], target_size[1]), some_brush) # Optional clear
                win32gui.DrawIconEx(hdc_mem, 0, 0, hicon, target_size[0], target_size[1], 0, 0, DI_NORMAL)

                # --- Prepare for GetDIBits ---
                class BITMAPINFOHEADER(Structure):
                    _fields_ = [('biSize', c_uint), ('biWidth', c_long), ('biHeight', c_long),
                                ('biPlanes', c_ushort), ('biBitCount', c_ushort), ('biCompression', c_uint),
                                ('biSizeImage', c_uint), ('biXPelsPerMeter', c_long), ('biYPelsPerMeter', c_long),
                                ('biClrUsed', c_uint), ('biClrImportant', c_uint)]

                class BITMAPINFO(Structure):
                     _fields_ = [('bmiHeader', BITMAPINFOHEADER), ('bmiColors', c_ubyte * (4 * 256))] # Placeholder for palette

                bmi = BITMAPINFO()
                bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER)
                bmi.bmiHeader.biWidth = target_size[0]
                bmi.bmiHeader.biHeight = -target_size[1] # Negative height for top-down DIB
                bmi.bmiHeader.biPlanes = 1
                bmi.bmiHeader.biBitCount = 32 # Request 32-bit BGRA
                bmi.bmiHeader.biCompression = win32con.BI_RGB # Standard uncompressed format

                buffer_size = target_size[0] * target_size[1] * 4 # 4 bytes/pixel
                buffer = (c_ubyte * buffer_size)()

                # --- Call GetDIBits with casted handle ---
                logging.debug(f"Attempting GetDIBits with hbmp type {type(hbmp)}, value {hbmp}")
                try:
                    # Explicitly cast hbmp to int
                    result = windll.gdi32.GetDIBits(hdc_mem, int(hbmp), 0, target_size[1], byref(buffer), byref(bmi), win32con.DIB_RGB_COLORS)
                except (TypeError, ctypes.ArgumentError) as e:
                    logging.error(f"GetDIBits failed during call. Error: {e}")
                    logging.error(f"Handle values: hdc_mem={hdc_mem}, hbmp={hbmp}")
                    return None # Exit if the call itself fails

                if result == 0:
                    error_code = win32api.GetLastError()
                    logging.error(f"GetDIBits returned 0 (failed). GetLastError(): {error_code}")
                    return None

                # --- Create PIL image from buffer ---
                # GetDIBits provides BGRA data when requesting 32bpp
                image = Image.frombuffer('RGBA', target_size, buffer, 'raw', 'BGRA', 0, 1)

                # Select the old bitmap back into the DC before deleting
                win32gui.SelectObject(hdc_mem, old_bmp)

                return image # Return the RGBA Image

            finally:
                # --- GDI Cleanup (Robust) ---
                if hbmp:
                    try: win32gui.DeleteObject(hbmp)
                    except win32gui.error as e: logging.debug(f"Error deleting hbmp: {e}")
                if hdc_mem:
                    try: win32gui.DeleteDC(hdc_mem)
                    except win32gui.error as e: logging.debug(f"Error deleting hdc_mem: {e}")
                if hdc:
                    try: win32gui.ReleaseDC(0, hdc) # Release screen DC
                    except win32gui.error as e: logging.debug(f"Error releasing hdc: {e}")
                if hicon:
                    try: win32gui.DestroyIcon(hicon)
                    except win32gui.error as e: logging.debug(f"Error destroying hicon: {e}")
                # Destroy any other handles returned by ExtractIconEx
                for handle in icons_to_destroy:
                    if handle and handle != hicon: # Avoid double-destroy
                        try: win32gui.DestroyIcon(handle)
                        except win32gui.error as e: logging.debug(f"Error destroying extra handle: {e}")

        else:
            logging.warning(f"Unsupported icon file type: {icon_path}")
            return None

    except UnidentifiedImageError:
        logging.error(f"Pillow cannot identify image file: {icon_path}")
        return None
    except FileNotFoundError:
        logging.error(f"Icon file not found during processing: {icon_path}")
        return None
    except Exception as e:
        logging.error(f"Unhandled error processing icon image {icon_path}: {e}")
        traceback.print_exc()
        return None


# --- Functions: get_dominant_color, create_colored_book_icon, set_shortcut_icon, load_backup, save_backup, apply_book_icons, revert_icons ---
# (Keep these as they were in the previous corrected version)
# Ensure set_shortcut_icon uses tuple: lnk.icon_location = (new_icon_path, 0)
# Ensure revert_icons parses string to tuple: lnk.icon_location = (path.strip(), index)
# --- Copy the rest of the functions (get_dominant_color, etc.) and the main block from the previous response ---
# --- Functions: get_dominant_color, create_colored_book_icon, set_shortcut_icon, load_backup, save_backup ---
# (These should remain largely the same as your improved version, ensure set_shortcut_icon uses tuple)
# --- Make sure set_shortcut_icon explicitly uses the tuple format ---
def set_shortcut_icon(shortcut_path, new_icon_path):
    """Applies the new icon file to the shortcut."""
    try:
        # Use tuple format (path, index) for icon_location
        new_icon_location = (new_icon_path, 0)
        with winshell.shortcut(shortcut_path) as lnk:
            lnk.icon_location = new_icon_location
        os.utime(shortcut_path, None) # Try to notify shell of change
        logging.info(f"Set icon for '{os.path.basename(shortcut_path)}' to '{os.path.basename(new_icon_path)}'")
        return True
    except Exception as e:
        logging.error(f"Failed to set icon for '{os.path.basename(shortcut_path)}': {e}")
        # traceback.print_exc() # Uncomment for detailed debugging
        return False

# --- Function: revert_icons ---
# (Your revert function looked correct in parsing the string back to tuple)
def revert_icons():
    """Reverts icons using the backup data."""
    logging.info("--- Starting Icon Reversion ---")
    backup_data = load_backup()

    if not backup_data:
        logging.warning("No backup data found or backup file is empty/corrupted.")
        print("No backup data found. Cannot revert.")
        return

    reverted_count = 0
    failed_count = 0

    print("Attempting to revert icons...")
    for shortcut_path, orig_icon_location_str in backup_data.items():
        shortcut_name = os.path.basename(shortcut_path)
        logging.info(f"Reverting {shortcut_name} to {orig_icon_location_str}")

        if not os.path.exists(shortcut_path):
             logging.warning(f"Shortcut path in backup does not exist anymore: {shortcut_path}")
             failed_count += 1
             continue

        try:
            with winshell.shortcut(shortcut_path) as lnk:
                # Parse the backup string "path,index" and set as tuple
                path = orig_icon_location_str
                index = 0
                if ',' in orig_icon_location_str:
                    try:
                        path, idx_str = orig_icon_location_str.rsplit(',', 1)
                        index = int(idx_str.strip())
                    except ValueError:
                         logging.error(f"Could not parse index from backup string: '{orig_icon_location_str}' for {shortcut_name}. Assuming index 0.")
                         index = 0 # Default index if parsing fails
                         path = orig_icon_location_str # Use the whole string as path if comma wasn't for index

                lnk.icon_location = (path.strip(), index) # Set as tuple

            os.utime(shortcut_path, None) # Notify shell
            logging.info(f"Successfully reverted icon for {shortcut_name}")
            reverted_count += 1
        except Exception as e:
            logging.error(f"Failed to revert icon for {shortcut_name}: {e}")
            # traceback.print_exc() # Uncomment for detailed debugging
            failed_count += 1

    logging.info(f"--- Reversion Finished: {reverted_count} reverted, {failed_count} failed. ---")
    print("\nFinished reverting icons.")
    print(f"  {reverted_count} icons reverted.")
    print(f"  {failed_count} icons failed (check log/console).")
    print("You might need to refresh your desktop (Right-click -> Refresh) or restart explorer.exe to see all changes.")


# --- REMAINING FUNCTIONS (copy from your version or the previous one) ---
def get_dominant_color(image):
    """Extracts the most dominant, suitable color from a PIL Image."""
    if image is None:
        return None
    try:
        max_dim = 128
        img_copy = image.copy()
        img_copy.thumbnail((max_dim, max_dim))
        buf = BytesIO()
        img_copy.save(buf, format='PNG')
        buf.seek(0)
        colors = colorgram.extract(buf, 10)
        if not colors:
            logging.warning("Colorgram could not extract any colors.")
            return None
        suitable_colors = []
        for c in colors:
            rgb = c.rgb
            rgb_tuple = (rgb.r, rgb.g, rgb.b)
            is_grayscale = abs(rgb.r - rgb.g) < 15 and abs(rgb.g - rgb.b) < 15
            is_too_dark = sum(rgb_tuple) < 50
            is_too_light = sum(rgb_tuple) > 700
            if not is_grayscale and not is_too_dark and not is_too_light:
                min_rgb, max_rgb = min(rgb_tuple), max(rgb_tuple)
                if max_rgb - min_rgb > 10:
                    suitable_colors.append((c, rgb_tuple))
        if suitable_colors:
            dominant_color = sorted(suitable_colors, key=lambda c: c[0].proportion, reverse=True)[0][1]
            logging.debug(f"Dominant suitable color: {dominant_color}")
            return dominant_color
        elif colors:
            c = sorted(colors, key=lambda c: c.proportion, reverse=True)[0]
            rgb_tuple = (c.rgb.r, c.rgb.g, c.rgb.b)
            logging.warning(f"No 'ideal' color found, using most prominent overall: {rgb_tuple}")
            return rgb_tuple
        else:
            logging.warning("No colors found after filtering.")
            return None
    except Exception as e:
        logging.error(f"Error getting dominant color: {e}")
        traceback.print_exc()
        return None

def create_colored_book_icon(target_color_rgb, output_icon_path):
    """Creates a new book icon with the cover colored and saves as .ico."""
    try:
        template = Image.open(TEMPLATE_PATH).convert("RGBA")
        target_color_rgba = (target_color_rgb[0], target_color_rgb[1], target_color_rgb[2], 255)
        modified_template = Image.new("RGBA", template.size)
        template_data = template.load()
        modified_data = modified_template.load()
        for y in range(template.height):
            for x in range(template.width):
                current_pixel = template_data[x, y]
                current_rgb = current_pixel[:3]
                if rgb_color_distance(current_rgb, TEMPLATE_COVER_COLOR_RGB) < COLOR_MATCH_TOLERANCE:
                    modified_data[x, y] = target_color_rgba[:3] + (current_pixel[3],)
                else:
                    modified_data[x, y] = current_pixel
        modified_template.save(output_icon_path, format='ICO', sizes=ICO_SIZES)
        logging.info(f"Saved new icon: {os.path.basename(output_icon_path)}")
        return True
    except FileNotFoundError:
        logging.error(f"Template icon not found at {TEMPLATE_PATH}")
        return False
    except Exception as e:
        logging.error(f"Error creating colored icon for {os.path.basename(output_icon_path)}: {e}")
        traceback.print_exc()
        return False


def load_backup():
    """Loads the backup data from the JSON file."""
    if os.path.exists(BACKUP_FILE):
        try:
            with open(BACKUP_FILE, 'r') as f:
                return json.load(f)
        except json.JSONDecodeError:
            logging.error(f"Backup file {BACKUP_FILE} is corrupted.")
            return {}
        except Exception as e:
            logging.error(f"Error reading backup file {BACKUP_FILE}: {e}")
            return {}
    return {}

def save_backup(backup_data):
    """Saves the backup data to the JSON file."""
    try:
        ensure_dir_exists(BACKUP_DIR)
        with open(BACKUP_FILE, 'w') as f:
            json.dump(backup_data, f, indent=4)
        logging.info(f"Backup saved to {BACKUP_FILE}")
    except Exception as e:
        logging.error(f"Error saving backup file: {e}")


def apply_book_icons():
    """Applies the book icons to desktop shortcuts."""
    logging.info("--- Starting Book Icon Application ---")
    ensure_dir_exists(GENERATED_ICONS_DIR)
    ensure_dir_exists(BACKUP_DIR)

    if not os.path.exists(TEMPLATE_PATH):
        logging.error(f"FATAL: Template icon '{os.path.basename(TEMPLATE_PATH)}' not found in script directory.")
        print(f"Error: Template icon '{os.path.basename(TEMPLATE_PATH)}' not found.")
        sys.exit(1)

    shortcuts = get_desktop_shortcuts()
    if not shortcuts:
        logging.warning("No shortcuts (.lnk files) found on the desktop.")
        print("No shortcuts found on the desktop.")
        return

    backup_data = load_backup()
    processed_count = 0
    skipped_count = 0

    for shortcut_path in shortcuts:
        shortcut_name = os.path.basename(shortcut_path)
        logging.info(f"\nProcessing shortcut: {shortcut_name}")

        # --- 1. Backup Original ---
        orig_icon_path, orig_icon_index = get_original_icon_info(shortcut_path)

        if orig_icon_path is None:
            logging.warning(f"Skipping {shortcut_name}: Could not get original icon info.")
            skipped_count += 1
            continue

        # Store backup if not already there or different
        # Ensure index is treated as integer for consistency in string format
        orig_icon_location_str = f"{orig_icon_path},{int(orig_icon_index)}"
        if shortcut_path not in backup_data or backup_data[shortcut_path] != orig_icon_location_str:
             backup_data[shortcut_path] = orig_icon_location_str
             logging.info(f"Backed up original icon: {orig_icon_location_str}")
        else:
             logging.info("Original icon already backed up correctly.")


        # --- 2. Extract Color ---
        safe_name = "".join(c for c in shortcut_name[:-4] if c.isalnum() or c in (' ', '_')).rstrip()
        generated_icon_filename = f"book_{safe_name}.ico"
        generated_icon_path = os.path.join(GENERATED_ICONS_DIR, generated_icon_filename)

        original_image = extract_icon_image(orig_icon_path, orig_icon_index)
        if original_image is None:
            logging.warning(f"Skipping {shortcut_name}: Could not extract image from original icon '{orig_icon_path}', index {orig_icon_index}.")
            skipped_count += 1
            continue

        dominant_color = get_dominant_color(original_image)
        if dominant_color is None:
            logging.warning(f"Skipping {shortcut_name}: Could not determine suitable dominant color.")
            skipped_count += 1
            continue

        logging.info(f"Dominant color found: RGB{dominant_color}")

        # --- 3. Create & Apply New Icon ---
        if create_colored_book_icon(dominant_color, generated_icon_path):
            if set_shortcut_icon(shortcut_path, generated_icon_path):
                 processed_count += 1
            else:
                 logging.error(f"Failed to apply new icon for {shortcut_name}, skipping.")
                 skipped_count += 1
        else:
            logging.error(f"Failed to create new icon for {shortcut_name}, skipping.")
            skipped_count += 1


    # --- 4. Final Backup Save ---
    save_backup(backup_data)
    logging.info(f"--- Finished: {processed_count} icons applied, {skipped_count} skipped. ---")
    print(f"\nFinished applying icons.")
    print(f"  {processed_count} icons successfully changed.")
    print(f"  {skipped_count} icons skipped (see log/console for reasons).")
    print(f"Original icon info backed up in: {BACKUP_DIR}")
    print("You might need to refresh your desktop (Right-click -> Refresh) or restart explorer.exe to see all changes.")


# --- Command Line Interface ---
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Replace desktop shortcut icons with color-matched book icons.")
    parser.add_argument(
        "action",
        choices=["apply", "revert"],
        help="'apply' to change icons to books, 'revert' to restore originals from backup."
    )

    args = parser.parse_args()

    print(f"Script directory: {SCRIPT_DIR}")
    print(f"Template path: {TEMPLATE_PATH}")
    print(f"Desktop path: {DESKTOP_PATH}")
    print(f"Backup directory: {BACKUP_DIR}")
    print(f"Generated icons directory: {GENERATED_ICONS_DIR}")


    if args.action == "apply":
        apply_book_icons()
    elif args.action == "revert":
        revert_icons()

    print("\nScript finished.")