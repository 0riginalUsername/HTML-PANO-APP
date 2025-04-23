import tkinter as tk
import os
import re
import piexif
import ftplib
import csv
import threading
import time
from dotenv import load_dotenv
from tkinter import filedialog, messagebox, ttk, simpledialog
from tkinter.simpledialog import askstring
from pathlib import Path
from jinja2 import Environment, FileSystemLoader
from datetime import datetime
from PIL import Image, ExifTags
from openpyxl import Workbook
from decimal import Decimal, getcontext
from concurrent.futures import ThreadPoolExecutor, as_completed
import smtplib
from email.message import EmailMessage

# -----------------------------------------------
# Gets .env file and extracts credentials
# -----------------------------------------------
env_file=Path(r"Z:\Survey\UT\_GabeA\PanoSandbox\.env")
load_dotenv(dotenv_path=env_file)
FTP_SERVER = os.getenv("FTP_SERVER")
FTP_USERNAME = os.getenv("FTP_USERNAME")
FTP_PASSWORD = os.getenv("FTP_PASSWORD")
EMAIL_HOST = os.getenv("EMAIL_HOST")
EMAIL_PORT = int(os.getenv("EMAIL_PORT", 587))
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
EMAIL_SENDER = os.getenv("EMAIL_SENDER")
# -----------------------------------------------
# Probe domain for current available connections
# -----------------------------------------------
def probe_ftp_max_threads(max_test=6):
    success_count = 0
    lock = threading.Lock()

    def try_connection():
        nonlocal success_count
        try:
            with ftplib.FTP(FTP_SERVER, timeout=10) as ftp:
                ftp.login(FTP_USERNAME, FTP_PASSWORD)
                with lock:
                    success_count += 1
        except Exception:
            pass  # Fail silently for test

    with ThreadPoolExecutor(max_workers=max_test) as executor:
        futures = [executor.submit(try_connection) for _ in range(max_test)]
        for future in as_completed(futures):
            pass

    print(f"[Thread Probe] Maximum simultaneous connections accepted: {success_count}")
    return max(1, success_count)  # fallback to at least 1

# ---------------------------------------------------------------------------
# Main function for folder selection and starting the processing workflow.
# ---------------------------------------------------------------------------
def choose_folder():
    # Show the warning window and wait for the user to close it.
        # Now prompt for the folder selection.
        while True:
            folder_path = filedialog.askdirectory(initialdir=".", title="Select a Folder")
            if folder_path:
                root.withdraw()  # Hide the main window after selection.
                break
            else:
                messagebox.showwarning("No Folder", "Please select a folder.")
        
        print(20 * ">", "Folder found! Selected folder:", folder_path)

        # Ask user for client and project names.
        client_name = askstring("Input", "Enter client name")
        if client_name:
            client_name = client_name.strip().replace(" ", "_")
        if not client_name:
            messagebox.showwarning("Missing Input", "Client name is required.")
            print(30 * "-", "Initialization failed: enter required input", 30 * "-")
            return

        project_name = askstring("Input", "Enter project name")
        if project_name:
            project_name = project_name.strip().replace(" ", "_")
        if not project_name:
            messagebox.showwarning("Missing Input", "Project name is required.")
            print(30 * "-", "Initialization failed: enter required input", 30 * "-")
            return
        
        employee_name = ask_user_name(["Allen", "Burt", "Gabe", "Kevin", "Morgan", "Nick", "Tanner"], root)

        # Create a date string for the project folder (e.g. "24Feb23").
        dt = datetime.now().strftime("%d%b%y")
        
        # Build the remote directory path.
        new_remote_dir = make_remote_domain_path(client_name, project_name, dt)
        # Process the selected folder to extract image metadata.
        images_dict = list_files_and_dirs(folder_path, new_remote_dir)
        
        # Compile the project data into HTML templates.
        proj_compile(client_name, folder_path, images_dict, new_remote_dir, project_name, employee_name)

# ------------------------------------------------------
# Ask user for their name that will appear in auto email
# ------------------------------------------------------

def ask_user_name(possible_names, root):
    result = {"name": None}
    # Name data asignment function
    def submit_name():
        selected = name_var.get()
        custom = custom_entry.get().strip()
        result["name"] = custom if custom else selected
        popup.destroy()
    # Tkinter windo logic
    popup = tk.Toplevel(root)
    popup.title("Select Your Name")
    popup.geometry("300x200")
    popup.grab_set()  # Makes this window modal (blocks interaction with others)

    tk.Label(popup, text="Choose your name:", font=("Arial", 12)).pack(pady=5)
    # Provides dropdown that iterates over names list possible_names
    name_var = tk.StringVar()
    name_combo = ttk.Combobox(popup, textvariable=name_var, values=possible_names)
    name_combo.pack(pady=5)
    name_combo.set(possible_names[0])
    # Second choice that allows custom name entry
    tk.Label(popup, text="Or enter your name:", font=("Arial", 10)).pack(pady=5)
    custom_entry = tk.Entry(popup)
    custom_entry.pack(pady=5)
    # Assigns name to variable
    submit_btn = tk.Button(popup, text="Submit", command=submit_name)
    submit_btn.pack(pady=10)

    root.wait_window(popup)  # Wait for the popup to close
    return result["name"]


# ---------------------------------------------------------------------------
# Constructs a remote directory path based on client, project, and date.
# Example output: "/auto/client_name/project_name/dt"
# ---------------------------------------------------------------------------
def make_domain_path(client_name, project_name, dt):
    new_dir = "/auto/" + client_name + "/" + project_name + "/" + dt
    return new_dir
# ---------------------------------------------------------------------------
# Constructs a remote directory path that is readable to local paths
# ---------------------------------------------------------------------------
def make_remote_domain_path(client_name, project_name, dt):
    new_dir = client_name + "/" + project_name + "/" + dt
    return new_dir

# ---------------------------------------------------------------------------
# Walks through the given folder and extracts JPEG image metadata using Pillow.
# Returns a dictionary mapping file names to their full path, formatted date, and base name.
# ---------------------------------------------------------------------------
def list_files_and_dirs(folder_path):
    image_files = {}
    file_count = 0

    print("Listing files in:", folder_path)
    # Iterate over all files in the selected folder.
    for f in os.listdir(folder_path):
        full_path = os.path.join(folder_path, f).replace("\\", "/")
        file_count += 1

        if os.path.isfile(full_path):
            print("  -", full_path)
            # Process only JPEG images.
            if f.lower().endswith(('.jpg', '.jpeg')):
                date_time = None
                format_date_time = None
                base_name = os.path.basename(full_path)
                try:
                    with Image.open(full_path) as img:
                        exif_data = img._getexif()
                        if exif_data is not None:
                            # Convert EXIF tag numbers to readable tag names.
                            exif = {ExifTags.TAGS.get(tag, tag): value for tag, value in exif_data.items()}
                            # Look for common date tags in the EXIF data.
                            date_time = (exif.get("DateTimeOriginal") or
                                         exif.get("DateTimeDigitized") or
                                         exif.get("DateTime"))
                            if date_time:
                                # Reformat the date string from "YYYY:MM:DD HH:MM:SS" to "YYYY/MM/DD HH:MM:SS"
                                format_date_time = re.sub(":", "/", date_time, count=2)
                                print("Found date for", f, ":", format_date_time)
                            else:
                                print("No date metadata found for", f)
                        else:
                            print("No EXIF data in", f)
                except Exception as e:
                    print("Error reading metadata from", full_path, ":", e)
                    format_date_time = None
                
                # Store the file information in the dictionary.
                image_files[f] = {
                    "full_path": full_path,
                    "date_time": format_date_time,
                    "base_name": base_name,
                }
    print(f"Total image files found: {file_count}")
    return image_files


# ---------------------------------------------------------------------------
# Compresses the image and saves it to a new directory on Z: drive based on client, project, and date.
# ---------------------------------------------------------------------------
def compress_image(input_image_path, remote_dir, quality=1):

    # Determine the output directory from remote_dir (expected format: "/auto/client/project/dt").
    base_output_directory = os.path.join("Z:/Survey/UT/ScriptFiles", remote_dir)
    compressed_dir = os.path.join(base_output_directory, "Compressed")
    os.makedirs(compressed_dir, exist_ok=True)

    try:
        # Load and sanitize EXIF data using piexif.
        try:
            exif_dict = piexif.load(input_image_path)

            # Remove the ColorSpace tag (40961) completely if present.
            if "0th" in exif_dict and 40961 in exif_dict["0th"]:
                print("Removing ColorSpace tag (40961) from EXIF")
                del exif_dict["0th"][40961]
            
            # Define other ICC-related tags to remove.
            icc_tags = {34675, 319, 318}
            # Remove these tags from all IFD groups except the thumbnail.
            for ifd in exif_dict:
                if ifd == "thumbnail":
                    continue
                for tag in list(exif_dict[ifd].keys()):
                    if tag in icc_tags:
                        print(f"Removing ICC-related tag {tag} from IFD '{ifd}'")
                        del exif_dict[ifd][tag]
            
            try:
                exif_bytes = piexif.dump(exif_dict)
            except Exception as e:
                print("Error dumping sanitized EXIF:", e)
                exif_bytes = None
        except Exception as e:
            print("Error loading or processing EXIF data:", e)
            exif_bytes = None

        # Open and process the image.
        with Image.open(input_image_path) as img:
            try:
                # Force conversion to RGB mode.
                img = img.convert("RGB")
            except Exception as e:
                print(f"Error converting image mode to RGB: {e}")
                return None

            # Remove any embedded ICC profile from the image info.
            if "icc_profile" in img.info:
                print("Stripping embedded ICC profile from image")
                del img.info["icc_profile"]

            new_path = os.path.join(compressed_dir, os.path.basename(input_image_path))

            try:
                # Attempt to save with the sanitized EXIF data.
                if exif_bytes:
                    img.save(new_path, "JPEG", quality=quality, exif=exif_bytes)
                else:
                    img.save(new_path, "JPEG", quality=quality)
                print("Image compressed and saved successfully:", new_path)
                return new_path
            except Exception as e:
                print(f"Error saving image with sanitized EXIF: {e}")
                # Fallback: save without EXIF data.
                try:
                    img.save(new_path, "JPEG", quality=quality)
                    print("Image compressed and saved without EXIF:", new_path)
                    return new_path
                except Exception as e2:
                    print(f"Final save attempt failed: {e2}")
                    return None

    except Exception as e:
        print(f"Error compressing image {input_image_path}: {e}")
        return None

# ---------------------------------------------------------------------------
# Renames images based on their date and compresses them. Returns a dictionary of renamed images.
# ---------------------------------------------------------------------------
def rename_images_by_date(images_dict, remote_dir, prefix="U"):
    images_list = []
    # Convert date strings to datetime objects and collect them
    for orig_name, info in images_dict.items():
        if info["date_time"]:
            try:
                dt = datetime.strptime(info["date_time"], "%Y/%m/%d %H:%M:%S")
            except Exception as e:
                dt = datetime.min
        else:
            dt = datetime.min
        images_list.append((orig_name, info, dt))
    # Sort images chronologically
    images_list.sort(key=lambda x: x[2])
    total_images = len(images_list)
    digits = max(2, len(str(total_images)))
    
    renamed_images = {}
    # Rename files and submit compression tasks in parallel
    max_workers = os.cpu_count() or 4
    with ThreadPoolExecutor(max_workers=(os.cpu_count() or 4)) as executor:
        futures = {}
        # Submit all compression tasks concurrently.
        for index, (orig_name, info, dt) in enumerate(images_list, start=1):
            new_number = f"{index:0{digits}d}"
            _, ext = os.path.splitext(orig_name)
            new_name = f"{prefix}{new_number}{ext.lower()}"
            
            old_path = info["full_path"]
            new_path = os.path.join(os.path.dirname(old_path), new_name)
            
            try:
                os.rename(old_path, new_path)
                # Submit the compression task.
                future = executor.submit(compress_image, new_path, remote_dir, quality=30)
                futures[future] = (new_name, info.copy(), new_path)
            except Exception as e:
                print(f"Error renaming or compressing {old_path}: {e}")
                info["compressed_path"] = None
                renamed_images[new_name] = info
        
        # Wait for all compression tasks to finish.
        for future in as_completed(futures):
            new_name, info, new_path = futures[future]
            try:
                compressed_path = future.result()
            except Exception as e:
                print(f"Error compressing image {new_name}: {e}")
                compressed_path = None
            
            info["full_path"] = new_path
            info["base_name"] = new_name 
            info["compressed_path"] = compressed_path
            renamed_images[new_name] = info
            
    return renamed_images

# ---------------------------------------------------------------------------
# Warns user about starting upload with internet connection
# ---------------------------------------------------------------------------
def show_connection_warning(renamed_images, remote_dir, proj_compiled, client_name, project_name, dt, employee_name, first_link):
    warn_win = tk.Toplevel(root)
    warn_win.title("Internet Connection Warning")
    warn_win.geometry("450x150")
    tk.Label(
        warn_win,
        text="ENSURE YOU HAVE A STABLE INTERNET CONNECTION\nDURING ENTIRE UPLOAD PROCESS",
        font=("Georgia", 9, "bold"),
        fg="red"
    ).pack(pady=20)
    tk.Button(
        warn_win,
        text="I understand",
        command=lambda: (warn_win.destroy(), show_compile_window(renamed_images, remote_dir, proj_compiled, employee_name, first_link))
    ).pack(pady=10)

# ---------------------------------------------------------------------------
# Starts upload function upon button press
# ---------------------------------------------------------------------------
def show_compile_window(renamed_images, remote_dir, proj_compiled, employee_name, first_link):
    compile_win = tk.Toplevel(root)
    compile_win.title("Project Compilation")
    compile_win.geometry("450x150")
    tk.Label(
        compile_win,
        text="Would you like to compile project data into an HTML template?",
        font=("Georgia", 10)
    ).pack(pady=30)
    tk.Button(
        compile_win,
        text="Yes!",
        command=lambda: (start_uploads(renamed_images, remote_dir, proj_compiled, employee_name, first_link), compile_win.destroy())
    ).pack(pady=10)


# ------------------------------------------------------------------------------------------
# Compiles project metadata (e.g., earliest date) and launches the next steps in processing.
# ------------------------------------------------------------------------------------------
def proj_compile(client_name, folder_path, images_dict, remote_dir, project_name, employee_name):
    # Gather all dates from image metadata.
    dates = []
    for info in images_dict.values():
        dt_str = info.get("date_time")
        if dt_str:
            try:
                dt = datetime.strptime(dt_str, "%Y/%m/%d %H:%M:%S")
                dates.append(dt)
            except Exception as e:
                print(f"Error parsing date for {info['full_path']}: {e}")
    
    if dates:
        dt_proj = min(dates)  # Use the earliest date.
        print("Derived project date from metadata:", dt_proj)
    else:
        dt_proj = datetime.now()
        print("No metadata date found; using current date:", dt_proj)
    
    # Format the project date for display.
    exif_date_str = dt_proj.strftime("%Y-%B-%d %H:%M")
    
    # Store compiled project information.
    proj_compiled = {
        "date_exif": exif_date_str,
        "name": client_name,
        "folder": folder_path,
        "proj_name": project_name
    }
    
    print("...Project compiled")
    print("Project summary:")
    print("Date:", exif_date_str, "\nName:", client_name, "\nFolder path:", folder_path)
    print("-" * 100)
    for file_name, info in images_dict.items():
        print(f"{file_name}: {info['full_path']}, Date: {info['date_time']}")
        print("=" * 100)
    
    # Rename images based on date.
    renamed_images = rename_images_by_date(images_dict, remote_dir, prefix="U")
    
    # Export GPS and date information to a CSV file, and returns the first completed link
    first_link = export_gps_and_date_to_csv(renamed_images, client_name, project_name)

    # Start window functions to initiate HTML upload
    show_connection_warning(renamed_images, remote_dir, proj_compiled, client_name, project_name, dt, employee_name, first_link)

# ------------------------------------------------------------------------------------------
# Renders HTML page using provided data and implements into page sourcecode
# ------------------------------------------------------------------------------------------
def render_template(file_name, info, proj_compiled, output_directory, template):
    # Convert the image's date string into a datetime object and reformat it.
    try:
        dt_obj = datetime.strptime(info["date_time"], "%Y/%m/%d %H:%M:%S")
    except Exception as e:
        print(f"Error parsing date for {file_name}: {e}")
        dt_obj = datetime.now()
    converted_dt = dt_obj.strftime("%d-%b-%y %I:%M:%S%p")
    # Inserts unique data and image file into page
    rendered_html = template.render(
        TITLE=proj_compiled["name"],
        DESCRIPTION=proj_compiled["date_exif"],
        IMG=info["base_name"],
        IMG_DATE=converted_dt,
    )
    base_name, _ = os.path.splitext(file_name)
    output_filename = os.path.join(output_directory, f"{base_name}.htm")
    with open(output_filename, "w", encoding="utf-8") as f:
        f.write(rendered_html)
    print(f"Created template: {output_filename}")
    return output_filename

# ------------------------------------------------------------------------------------------
# Renders email HTML code with relevent data, and sends email to listed reciepients
# ------------------------------------------------------------------------------------------
def send_html_email(project_name, client_name, date, employee_name, first_link, remote_dir):
    # Gets template location and creates HTML page
    template_path = Path(r"Z:\Survey\UT\_GabeA\PanoSandbox\Proj\Email-Report-Template.htm")
    env = Environment(loader=FileSystemLoader(template_path.parent))
    template = env.get_template(template_path.name)
    # Inserts relevent pano information into email HTML page
    html_content = template.render(
        PROJECT_NAME=str(project_name[0]) if isinstance(project_name, tuple) else project_name,
        CLIENT_NAME=str(client_name[0]) if isinstance(client_name, tuple) else client_name,
        UPLOAD_TIME=str(date[0]) if isinstance(date, tuple) else date,
        EMPLOYEE=str(employee_name[0]) if isinstance(employee_name, tuple) else employee_name,
        PANO_LINK=str(first_link[0]) if isinstance(first_link, tuple) else first_link,
        DIRECTORY_PATH=Path(get_local_directory(remote_dir)).as_posix()
    )

    # Sends email with provided crendentials
    msg = EmailMessage()
    msg["Subject"] = f"✅ Upload Complete - {str(project_name[0]) if isinstance(project_name, tuple) else project_name}"
    msg["From"] = EMAIL_SENDER
    msg["To"] = ",".join([""
    "gabe.alley@sunrise-eng.com",
    "kdawson@sunrise-eng.com"
   ])
    msg.set_content("Your upload is complete.")
    msg.add_alternative(html_content, subtype="html")

    try:
        with smtplib.SMTP(EMAIL_HOST, 587) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(EMAIL_USER, "Sunrise2019")
            smtp.send_message(msg)
            print(">>Email successfully sent!<<")
    except Exception as e:
            print("❌ Failed to send email via iPage:", e)

# ---------------------------------------------------------------------------
# Creates HTML templates from project data and images, saving them to a remote directory on Z: drive.
# ---------------------------------------------------------------------------
def make_proj_template(proj_compiled, images_dict, remote_dir):
    # Creates the final directory of pano backup files
    output_directory = os.path.join("Z:/Survey/UT/ScriptFiles", remote_dir)
  
    
    # Create the output directory if it doesn't exist.
    os.makedirs(output_directory, exist_ok=True)
    print("Output directory created:", output_directory)
    
    # Load the HTML template.
    env = Environment(loader=FileSystemLoader(r"Z:\Survey\UT\_GabeA\PanoSandbox\Proj"))
    try:
        template = env.get_template("Pano-Template.htm")
    except Exception as e:
        print("Error loading template:", e)
        return []
    
    html_files = []
    # Use a ThreadPoolExecutor to render and save templates concurrently.
    with ThreadPoolExecutor(max_workers=(os.cpu_count() or 4)) as executor:
        futures = {}
        for file_name, info in images_dict.items():
            future = executor.submit(render_template, file_name, info, proj_compiled, output_directory, template)
            futures[future] = file_name
        for future in as_completed(futures):
            try:
                result = future.result()
                html_files.append(result)
            except Exception as e:
                print(f"Error rendering template for {futures[future]}: {e}")
    
    return html_files


# ---------------------------------------------------------------------------
# Helper function to convert GPS coordinates (from EXIF) to decimal degrees.
# Expects a tuple of three tuples (rational format).
# ---------------------------------------------------------------------------
def convert_to_degrees_with_ref(value, ref):
        # sets up precise calculation for decimal degrees conversion
        getcontext().prec = 28  
        try:
            if isinstance(value, (list, tuple)) and len(value) == 3:
                if isinstance(value[0], tuple):
                    d = value[0][0] / value[0][1]
                    m = value[1][0] / value[1][1]
                    s = value[2][0] / value[2][1]
                else:
                    d, m, s = value
                result = d + (m / 60.0) + (s / 3600.0)
            else:
                result = float(value)
            if isinstance(ref, str) and ref.upper() in ["S", "W"]:
                result = -result
            elif isinstance(ref, (int, float)) and ref == 1:
                result = -result
            return result
        except Exception as e:
            print(f"Error converting {value} with ref {ref}: {e}")
            return None


def upload_file_via_ftp(file_path, remote_dir, max_retries=3, delay_base=2):
    """
    Upload a file to the FTP server, with optional retry support.
    """
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return

    for attempt in range(1, max_retries + 1):
        try:
            print(f"Starting upload for: {file_path} (Attempt {attempt})")
            with ftplib.FTP(FTP_SERVER, timeout=30) as ftp:
                ftp.login(FTP_USERNAME, FTP_PASSWORD)
                ftp.cwd(remote_dir)
                with open(file_path, 'rb') as f:
                    ftp.storbinary(f"STOR {os.path.basename(file_path)}", f)
            print(f"Successfully uploaded {os.path.basename(file_path)} on attempt {attempt}")
            return  # Success, exit the loop
        except Exception as e:
            print(f"[{threading.current_thread().name}] Attempt {attempt} failed for {file_path}: {e}")
            if attempt == max_retries:
                print(f"Giving up on {file_path} after {max_retries} attempts.")
            else:
                sleep_time = delay_base ** attempt
                print(f"Retrying in {sleep_time} seconds...")
                time.sleep(sleep_time)



# ---------------------------------------------------------------------------
# Recursively creates directories on the FTP server as needed.
# ---------------------------------------------------------------------------
def create_remote_directory_recursive(ftp, remote_dir):
    """
    Ensures that the remote directory exists by attempting to navigate to it or creating it recursively.
    """
    try:
        ftp.cwd(remote_dir)
        return True
    except ftplib.error_perm:
        # Attempt to create directories recursively.
        dirs = remote_dir.split('/')
        current_dir = ''
        for d in dirs:
            if d:  # skip empty parts
                current_dir += '/' + d
                try:
                    ftp.mkd(current_dir)
                except ftplib.error_perm:
                    # If directory already exists, just continue.
                    pass
        try:
            ftp.cwd(remote_dir)
            return True
        except Exception:
            return False


# ---------------------------------------------------------------------------
# Uploads HTML templates to the remote directory via FTP.
# ---------------------------------------------------------------------------
def upload_html_templates_concurrently(html_files, remote_dir):

    # Print the remote directory being used.
    print("Remote directory to use:", remote_dir)
    
    # First, check (and/or create) the remote directory using a temporary FTP connection.
    try:
        with ftplib.FTP(FTP_SERVER, timeout=30) as ftp:
            ftp.login(FTP_USERNAME, FTP_PASSWORD)
            if not create_remote_directory_recursive(ftp, remote_dir):
                print(f"[{threading.current_thread().name}] Could not create remote directory for {remote_dir}.")
                return
    except Exception as e:
        print(f"Error establishing FTP connection: {e}")
        return

    max_threads = probe_ftp_max_threads()

    # Now, for each HTML file, submit an upload task that creates its own FTP connection.
    with ThreadPoolExecutor(max_workers=max_threads) as executor:
        for file in html_files:
            # file is the full path to the HTML template.
            executor.submit(upload_file_via_ftp_with_retry, file, remote_dir)



# Helper function to get the local directory based on remote_dir.
def get_local_directory(remote_dir):
    local_dir = os.path.join("Z:/Survey/UT/ScriptFiles", remote_dir)
    return local_dir




def upload_images_concurrently(image_files, remote_dir):
    max_threads = probe_ftp_max_threads()

    with ThreadPoolExecutor(max_workers=max_threads) as executor:
        for filename, file_info in image_files.items():
            try:
                executor.submit(upload_file_via_ftp_with_retry, file_info['compressed_path'], remote_dir)
            except Exception as e:
                print(f"[{threading.current_thread().name}] Error uploading {filename}: {e}")




# ---------------------------------------------------------------------------
# Extracts the Date Taken from an image using common EXIF tags.
# Returns a formatted date string.
# ---------------------------------------------------------------------------
def extract_date_taken(image_path):
    try:
        with Image.open(image_path) as img:
            exif_data = img._getexif()
            if exif_data is not None:
                exif = {ExifTags.TAGS.get(tag, tag): value for tag, value in exif_data.items()}
                date_taken = (exif.get("DateTimeOriginal") or
                              exif.get("DateTimeDigitized") or
                              exif.get("DateTime"))
                # Reformat from "YYYY:MM:DD HH:MM:SS" to "YYYY/MM/DD HH:MM:SS"
                format_date_time = re.sub(":", "/", date_taken, count=2)
                return format_date_time
    except Exception as e:
        print(f"Error extracting date from {image_path}: {e}")
    return None


# ---------------------------------------------------------------------------
# Starts the upload process (both HTML templates and images) in a background thread.
# ---------------------------------------------------------------------------
def start_uploads(renamed_images, remote_dir, proj_compiled, employee_name, first_link):

    def run_uploads():
        html_files = make_proj_template(proj_compiled, renamed_images, remote_dir)
        upload_html_templates_concurrently(html_files, remote_dir)
        upload_images_concurrently(renamed_images, remote_dir)
        show_upload_complete_window(remote_dir, employee_name, proj_compiled, first_link)

    threading.Thread(target=run_uploads, daemon=True).start()

def show_upload_complete_window(remote_dir, employee_name, proj_compiled, first_link):

    CLIENT_NAME=proj_compiled["name"],
    PROJ_NAME=proj_compiled["proj_name"],
    DATE=datetime.now().strftime("%Y-%m-%d %I:%M %p")
    send_html_email(PROJ_NAME, CLIENT_NAME, DATE, employee_name, first_link, remote_dir)
    # Determine the local directory from remote_dir.
    local_dir = get_local_directory(remote_dir)
    
    # Create a new window.
    complete_win = tk.Toplevel(root)
    complete_win.title("Upload Complete")
    complete_win.geometry("300x150")
    
    # Display a completion message.
    tk.Label(complete_win, text="Upload complete!", font=("Georgia", 12)).pack(pady=10)
    
    # Button to open the directory.
    tk.Button(
        complete_win, 
        text="Open Directory", 
        command=lambda: os.startfile(local_dir)
    ).pack(pady=10)


# ---------------------------------------------------------------------------
# Main export function for CSV.
# Walks through the folder, extracts GPS and date data from JPEGs,
# converts GPS values, and writes everything (plus a hyperlink) to a CSV file.
# The CSV is saved in:
# Z:/Survey/UT/ScriptFiles/<client_name>/<project_name>/<dt>/
# ---------------------------------------------------------------------------
def export_gps_and_date_to_csv(renamed_images, client_name, project_name):

    # Sanitize client and project names.
    client_name = client_name.strip()
    project_name = project_name.strip()

    
    # Build the output CSV filename.
    current_time = datetime.now().strftime("%d%b%y")
    
    output_csv = f"{current_time}_{client_name}_{project_name}.csv"

    # Build the output directory on Z: drive including the date folder.
    output_directory = os.path.join("Z:/Survey/UT/ScriptFiles", client_name, project_name, current_time)
    os.makedirs(output_directory, exist_ok=True)
    
    # Build the full path to the CSV file.
    output_file_path = os.path.join(output_directory, output_csv)
    print("Saving CSV to:", output_file_path)
    

    # Extracts gps Lat Long from images
    def extract_gps_data(image_path):
        try:
            with Image.open(image_path) as img:
                exif_data = img._getexif()
                if not exif_data:
                    return None
                exif = {ExifTags.TAGS.get(tag, tag): value for tag, value in exif_data.items()}
                gps_info = exif.get("GPSInfo")
                if not gps_info:
                    return None
                gps_data = {ExifTags.GPSTAGS.get(key, key): value for key, value in gps_info.items()}
                return gps_data
        except Exception as e:
            print(f"Error extracting GPS data from {image_path}: {e}")
            return None
    # Extracts date/time from images
    def extract_date_taken(image_path):
        try:
            with Image.open(image_path) as img:
                exif_data = img._getexif()
                if exif_data is not None:
                    exif = {ExifTags.TAGS.get(tag, tag): value for tag, value in exif_data.items()}
                    date_taken = (exif.get("DateTimeOriginal") or
                                  exif.get("DateTimeDigitized") or
                                  exif.get("DateTime"))
                    # Reformat from "YYYY:MM:DD HH:MM:SS" to "YYYY/MM/DD HH:MM:SS"
                    format_date_time = re.sub(":", "/", date_taken, count=2)
                    return format_date_time
        except Exception as e:
            print(f"Error extracting date from {image_path}: {e}")
        return None
    # Sets up first_hyperlink variable
    first_hyperlink = None  
    # takes all data extracted and imports into .csv format
    with open(output_file_path, "w", newline="", encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["Filename", "Date Taken", "GPSLatitude", "GPSLongitude", "GPSAltitude", "Hyperlink"])
        domain_path = make_domain_path(client_name, project_name, current_time)
        
        for idx, (file_name, info) in enumerate(renamed_images.items()):
            base_name, _ = os.path.splitext(info["base_name"])
            html_filename = base_name + ".htm"
            hyperlink = "https://www.seihds.com" + domain_path + "/" + html_filename

            if idx == 0:
                first_hyperlink = hyperlink  

            date_taken = info.get("date_time") or extract_date_taken(info["full_path"])
            gps_data = extract_gps_data(info["full_path"])

            if gps_data:
                lat_raw = gps_data.get("GPSLatitude")
                lat_ref = gps_data.get("GPSLatitudeRef")
                lat = convert_to_degrees_with_ref(lat_raw, lat_ref) if lat_raw and lat_ref else None
                lon_raw = gps_data.get("GPSLongitude")
                lon_ref = gps_data.get("GPSLongitudeRef")
                lon = convert_to_degrees_with_ref(lon_raw, lon_ref) if lon_raw and lon_ref else None
                alt_raw = gps_data.get("GPSAltitude")
                alt_ref = gps_data.get("GPSAltitudeRef")
                alt = convert_to_degrees_with_ref(alt_raw, alt_ref) if alt_raw and alt_ref else None
            else:
                lat, lon, alt = (None, None, None)

            writer.writerow([info["base_name"], date_taken, lat, lon, alt, hyperlink])
    print("CSV file saved as", output_file_path)
    return first_hyperlink


# ---------------------------------------------------------------------------
# Initialize the main Tkinter window and UI elements.
# ---------------------------------------------------------------------------
root = tk.Tk()
root.geometry("600x400") 
root.title("Pano image process")
# Warning text above the folder prompt.
tk.Label(root, 
         text="ENSURE YOU HAVE GONE THROUGH ALL PANO PHOTOS\nAND DELETED UNWANTED OR DUPLICATE PANOS",
         font=("Georgia", 12, "bold"),
         fg="red").pack(pady=10)

# Existing prompt for folder selection.
tk.Label(root, 
         text="Find your pano image folder in directory:", 
         font=("Georgia", 15, "bold")).pack(pady=30)

button = tk.Button(root, text="Choose Folder", command=choose_folder)
button.pack(pady=40)

root.mainloop()