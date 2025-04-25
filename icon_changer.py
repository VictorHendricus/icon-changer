import os
import json
import winshell
import sys
from PIL import Image, UnidentifiedImageError
import colorgram
import argparse
import traceback
import logging
import win32gui  # type: ignore
from PIL import ImageWin  # type: ignore
# Use literal for DI_NORMAL to avoid win32con import
DI_NORMAL = 0x0003

# --- Configuration ---
# Color of the book cover in the template image (adjust if needed)
# Use a color picker tool on 'book_template.png' to get the exact RGB
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

# --- Helper Functions ---

def ensure_dir_exists(dir_path):
    """Creates a directory if it doesn't exist."""
    os.makedirs(dir_path, exist_ok=True)

def rgb_color_distance(color1, color2):
    """Calculates the 'distance' between two RGB colors."""
    return sum([(c1 - c2) ** 2 for c1, c2 in zip(color1, color2)]) ** 0.5

def get_desktop_shortcuts():
    """Finds all .lnk files on the desktop."""
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

def get_original_icon_info(shortcut_path):
    """Gets the original icon location (path, index) from a shortcut."""
    try:
        with winshell.shortcut(shortcut_path) as lnk:
            icon_location = lnk.icon_location
            # icon_location can be a string or tuple
            if isinstance(icon_location, tuple):
                path = icon_location[0]
                index = icon_location[1] if len(icon_location) > 1 else 0
            elif isinstance(icon_location, str):
                parts = icon_location.split(',')
                path = parts[0].strip()
                index = int(parts[1].strip()) if len(parts) > 1 else 0
            else:
                path = None
                index = 0
            if not path:
                # If no specific icon is set, it might use the target's default
                target_path = lnk.path
                if target_path and os.path.exists(target_path):
                    logging.info(f"No explicit icon set for {os.path.basename(shortcut_path)}, using target: {target_path}")
                    return target_path, 0
                else:
                    logging.warning(f"Cannot determine icon for {os.path.basename(shortcut_path)} - no icon_location and invalid target.")
                    return None, None
            # Resolve environment variables in path (like %SystemRoot%)
            path = os.path.expandvars(path)
            return path, index
    except Exception as e:
        logging.error(f"Error reading shortcut '{os.path.basename(shortcut_path)}': {e}")
        return None, None

def extract_icon_image(icon_path, index):
    """
    Attempts to extract the icon as a PIL Image object.
    Uses win32gui.ExtractIconEx and PIL.ImageWin.Dib for exe/dll robustly.
    """
    if not icon_path or not os.path.exists(icon_path):
        logging.warning(f"Icon path does not exist: {icon_path}")
        return None
    try:
        if icon_path.lower().endswith(".ico"):
            img = Image.open(icon_path)
            img.load() # Load image data
            if img.mode != 'RGBA':
                img = img.convert('RGBA')
            return img
        elif icon_path.lower().endswith((".exe", ".dll")):
            logging.warning(f"Attempting to extract icon from {os.path.basename(icon_path)}. Quality may vary or fail.")
            try:
                large, small = win32gui.ExtractIconEx(icon_path, index, 1)
                hicon = large[0] if large else None
                if not hicon:
                    logging.error(f"No icon extracted from {icon_path} at index {index}")
                    return None
                # Get icon size (use 256x256 if available, else 32x32)
                icon_w, icon_h = 256, 256
                # Create a device context and bitmap
                hdc = win32gui.GetDC(0)
                hbmp = win32gui.CreateCompatibleBitmap(hdc, icon_w, icon_h)
                hdc_mem = win32gui.CreateCompatibleDC(hdc)
                win32gui.SelectObject(hdc_mem, hbmp)
                win32gui.DrawIconEx(hdc_mem, 0, 0, hicon, icon_w, icon_h, 0, 0, DI_NORMAL)
                dib = ImageWin.Dib('RGB', (icon_w, icon_h))
                dib.frombytes(win32gui.GetBitmapBits(hbmp, True))
                img = dib.tobytes()
                image = Image.frombytes('RGB', (icon_w, icon_h), img)
                image = image.convert('RGBA')
                # Cleanup
                win32gui.DestroyIcon(hicon)
                win32gui.DeleteObject(hbmp)
                win32gui.DeleteDC(hdc_mem)
                win32gui.ReleaseDC(0, hdc)
                return image
            except Exception as e:
                logging.error(f"win32gui icon extraction failed for {icon_path}: {e}")
                return None
        else:
            logging.warning(f"Unsupported icon file type: {icon_path}")
            return None
    except UnidentifiedImageError:
        logging.error(f"Cannot identify image file (maybe corrupt or unsupported format?): {icon_path}")
        return None
    except FileNotFoundError:
        logging.error(f"Icon file not found: {icon_path}")
        return None
    except Exception as e:
        logging.error(f"Error loading icon image {icon_path}: {e}")
        return None


def get_dominant_color(image):
    """Extracts the most dominant, suitable color from a PIL Image."""
    if image is None:
        return None
    try:
        # Resize for faster processing, preserving aspect ratio
        max_dim = 128
        image.thumbnail((max_dim, max_dim))

        # Extract colors using colorgram
        # Extract more colors initially to have options
        colors = colorgram.extract(image, 10)

        if not colors:
            logging.warning("Colorgram could not extract any colors.")
            return None

        # Filter out colors that are too light, too dark, or too transparent
        suitable_colors = []
        for c in colors:
            rgb = c.rgb
            # Check saturation/lightness thresholds (simple approach)
            is_grayscale = abs(rgb.r - rgb.g) < 15 and abs(rgb.g - rgb.b) < 15
            is_too_dark = sum(rgb) < 50
            is_too_light = sum(rgb) > 700 # Max is 765 (255*3)

            # Colorgram doesn't directly give alpha, but often skips transparent areas.
            # If the source image had alpha, we assume colorgram picked opaque parts.

            if not is_grayscale and not is_too_dark and not is_too_light:
                suitable_colors.append(c)

        # Choose the most prominent suitable color, or fallback to the overall prominent
        if suitable_colors:
            dominant_color = sorted(suitable_colors, key=lambda c: c.proportion, reverse=True)[0]
            logging.debug(f"Dominant suitable color: {dominant_color.rgb}")
            return dominant_color.rgb
        elif colors:
             # Fallback: use the most prominent color overall if no "suitable" ones found
             dominant_color = sorted(colors, key=lambda c: c.proportion, reverse=True)[0]
             logging.warning(f"No 'ideal' color found, using most prominent: {dominant_color.rgb}")
             return dominant_color.rgb
        else:
            # Should not happen if colors were extracted, but as a safeguard
            return None

    except Exception as e:
        logging.error(f"Error getting dominant color: {e}")
        # traceback.print_exc() # Uncomment for detailed debugging
        return None


def create_colored_book_icon(target_color_rgb, output_icon_path):
    """Creates a new book icon with the cover colored and saves as .ico."""
    try:
        template = Image.open(TEMPLATE_PATH).convert("RGBA")
        target_color_rgba = (target_color_rgb.r, target_color_rgb.g, target_color_rgb.b, 255) # Ensure full opacity

        modified_template = Image.new("RGBA", template.size)

        for x in range(template.width):
            for y in range(template.height):
                current_pixel = template.getpixel((x, y))
                current_rgb = current_pixel[:3]
                current_alpha = current_pixel[3]

                # Check if the pixel color is close to the template's cover color
                if rgb_color_distance(current_rgb, TEMPLATE_COVER_COLOR_RGB) < COLOR_MATCH_TOLERANCE:
                    # Replace with target color, keeping original alpha if needed (though usually cover is opaque)
                    # Use target_color_rgba directly if you want the new cover fully opaque
                    # Or blend: new_pixel = target_color_rgba[:3] + (current_alpha,)
                    modified_template.putpixel((x, y), target_color_rgba)
                else:
                    # Keep the original pixel
                    modified_template.putpixel((x, y), current_pixel)

        # Save as ICO with multiple sizes
        modified_template.save(output_icon_path, format='ICO', sizes=ICO_SIZES)
        logging.info(f"Saved new icon: {os.path.basename(output_icon_path)}")
        return True

    except FileNotFoundError:
         logging.error(f"Template icon not found at {TEMPLATE_PATH}")
         return False
    except Exception as e:
        logging.error(f"Error creating colored icon for {os.path.basename(output_icon_path)}: {e}")
        # traceback.print_exc() # Uncomment for detailed debugging
        return False


def set_shortcut_icon(shortcut_path, new_icon_path):
    """Applies the new icon file to the shortcut."""
    try:
        # For .ico files, the index is typically 0
        with winshell.shortcut(shortcut_path) as lnk:
            lnk.icon_location = (new_icon_path, 0)
        os.utime(shortcut_path, None)
        logging.info(f"Set icon for '{os.path.basename(shortcut_path)}' to '{os.path.basename(new_icon_path)}'")
        return True
    except Exception as e:
        logging.error(f"Failed to set icon for '{os.path.basename(shortcut_path)}': {e}")
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


# --- Main Actions ---

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
        orig_icon_location_str = f"{orig_icon_path},{orig_icon_index}"
        if shortcut_path not in backup_data or backup_data[shortcut_path] != orig_icon_location_str:
             backup_data[shortcut_path] = orig_icon_location_str
             logging.info(f"Backed up original icon: {orig_icon_location_str}")
        else:
             logging.info("Original icon already backed up correctly.")


        # --- 2. Extract Color ---
        # Clean shortcut name for use in filename
        safe_name = "".join(c for c in shortcut_name[:-4] if c.isalnum() or c in (' ', '_')).rstrip() # Remove .lnk extension
        generated_icon_filename = f"book_{safe_name}.ico"
        generated_icon_path = os.path.join(GENERATED_ICONS_DIR, generated_icon_filename)

        # Check if we already generated this icon (useful if run multiple times)
        # Optional: Add a check here to skip regeneration if file exists and source hasn't changed?

        original_image = extract_icon_image(orig_icon_path, orig_icon_index)
        if original_image is None:
            logging.warning(f"Skipping {shortcut_name}: Could not extract image from original icon path '{orig_icon_path}'.")
            skipped_count += 1
            continue # Skip to next shortcut

        dominant_color = get_dominant_color(original_image)
        if dominant_color is None:
            logging.warning(f"Skipping {shortcut_name}: Could not determine dominant color.")
            skipped_count += 1
            continue # Skip to next shortcut

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
            # Use winshell directly to set the original location string
            with winshell.shortcut(shortcut_path) as lnk:
                # Parse the backup string "path,index" and set as tuple
                if ',' in orig_icon_location_str:
                    path, idx = orig_icon_location_str.rsplit(',', 1)
                    lnk.icon_location = (path.strip(), int(idx.strip()))
                else:
                    lnk.icon_location = (orig_icon_location_str.strip(), 0)
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