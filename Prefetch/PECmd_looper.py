# Python 3 Script used for testing behaviour in Microsoft Windows Prefetch files.
# Please note that script uses tool PECmd created by Eric Zimmerman (saericzimmerman@gmail.com) Source:
# https://github.com/EricZimmerman/PECmd
#
#
# WARNING: Option 9 can crush Google Chrome
# Script can be improved but for testing purposes, in this case it was enough

import os
import subprocess
import shutil
import psutil
import logging
import json
import time
from pprint import pprint

# Configure logging
logging.basicConfig(
    filename="PECmd_looper.log", 
    level=logging.INFO, 
    format="%(asctime)s - %(levelname)s - %(message)s"
)




def start_and_close_program(program_path,program_name):
    """
    Starts Google program from the specified path, waits for 15 seconds, then closes it.

    :param program_path: Full path to the Google program executable.
    """
    try:
        # Start program
        process = subprocess.Popen(program_path, shell=True)
        logging.info(f"Started program: {program_path}")
        print(f"Started program: {program_path}")

        # Wait 15 seconds
        time.sleep(30)

        # Find and close program processes
        closed = False
        for proc in psutil.process_iter(attrs=['pid', 'name']):
            if program_name in proc.info['name'].lower():
                proc.terminate()  # Terminate program
                closed = True

        if closed:
            logging.info("Closed Google program after 15 seconds.")
            print("Closed Google program.")
        else:
            logging.warning("No running program process found to close.")
            print("No running program process found to close.")

    except Exception as e:
        logging.error(f"Error: {e}")
        print(f"Error: {e}")




def copy_prefetch_files(source_folder, destination_folder, prefix="", suffix=""):
    """
    Copies Prefetch files from the source folder to the destination folder, 
    filtering by specified prefix and suffix.

    :param source_folder: Folder where Prefetch files are located.
    :param destination_folder: Folder where filtered Prefetch files will be copied.
    :param prefix: (Optional) Prefix to filter files.
    :param suffix: (Optional) Suffix to filter files.
    """
    if not os.path.exists(source_folder):
        logging.error(f"Source folder not found: {source_folder}")
        print(f"Error: Source folder not found: {source_folder}")
        return

    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)  # Create destination if it doesn't exist
        logging.info(f"Created destination folder: {destination_folder}")

    # Collect matching Prefetch files
    matching_files = [
        file for file in os.listdir(source_folder)
        if file.startswith(prefix) and file.endswith(suffix)
    ]

    if not matching_files:
        logging.warning("No matching Prefetch files found.")
        print("No matching Prefetch files found.")
        return

    # Copy matching files
    for file in matching_files:
        source_path = os.path.join(source_folder, file)
        destination_path = os.path.join(destination_folder, file)

        try:
            shutil.copy2(source_path, destination_path)  # Copy with metadata
            logging.info(f"Copied: {source_path} -> {destination_path}")
            print(f"Copied: {file} -> {destination_folder}")

        except Exception as e:
            logging.error(f"Error copying {file}: {e}")
            print(f"Error copying {file}: {e}")






def process_files(process_search, folder_path, command_template):
    """
    Finds all files in the specified folder that match the prefix 'MSEDGE.EXE-' 
    and executes a given command using cmd.exe for each file.
    
    :param folder_path: The directory to scan for matching files.
    :param command_template: The command template where '{}' will be replaced by the file path.
    """
    prefix,suffix = process_search

    if not os.path.exists(folder_path):
        logging.error(f"Folder not found: {folder_path}")
        return
    
    # Collect matching file paths
    file_paths = [
        os.path.join(folder_path, file)
        for file in os.listdir(folder_path)
            if file.startswith(prefix) and file.endswith(suffix)
    ]

    if not file_paths:
        logging.warning("No matching files found.")
        return
    
    logging.info(f"Found {len(file_paths)} matching files.")

    # Iterate over files and execute the command
    for file_path in file_paths:

        filename = os.path.splitext(os.path.basename(file_path))[0]

        command = command_template.format(file_path,filename)
        logging.info(f"Executing: {command}")

        try:
            subprocess.run(command, shell=True, check=True)
            logging.info(f"Successfully processed: {file_path}")
        except subprocess.CalledProcessError as e:
            logging.error(f"Error processing {file_path}: {e}")







def process_json_files(analysis_search,folder_path,counter):
    """
    Reads JSON files in the given folder that match the naming pattern "202501241*_PECmd_Output.json",
    extracts "ExecutableName" and "Hash", and prints them in the format "ExecutableName-Hash"
    along with "LastRun" and "RunCount".
    
    :param folder_path: The directory to scan for JSON files.
    """
    prefix,suffix = analysis_search

    if not os.path.exists(folder_path):
        logging.error(f"Folder not found: {folder_path}")
        return
    
    # Collect matching JSON file paths
    json_files = [
        os.path.join(folder_path, file)
        for file in os.listdir(folder_path)
        if file.startswith(prefix) and file.endswith(suffix)
    ]

    if not json_files:
        logging.warning("No matching JSON files found.")
        return
    
    logging.info(f"Found {len(json_files)} matching JSON files.")
    logging.warning(f"### RUN {counter} ###")
    # Iterate through JSON files
    for json_file in json_files:
        try:
            with open(json_file, "r", encoding="utf-8") as f:
                data = json.load(f)  # Load JSON content

            # Extract necessary fields
            executable_name = data.get("ExecutableName", "Unknown")
            hash_value = data.get("Hash", "NoHash")
            last_run = data.get("LastRun", "Unknown")
            run_count = data.get("RunCount", "Unknown")
            run_count = data.get("RunCount", "Unknown")

            directories = data.get("Directories", "Unknown").split(",")
            directories_count =len(directories)
            filesloaded = data.get("FilesLoaded", "Unknown").split(",")
            filesloaded_count =len(filesloaded)
            

            # Format and print output
            formatted_output = f"{executable_name}-{hash_value}"
            logging.warning(f"{formatted_output} | Last Run: {last_run} | Run Count: {run_count} | Dir Count: {directories_count} | File Count: {filesloaded_count}")

            #logging.info(f"Processed: {json_file}")

        except (json.JSONDecodeError, FileNotFoundError, KeyError) as e:
            logging.error(f"Error processing {json_file}: {e}")







def view_json_files(folder_path,analysis_search):
    """
    Lists JSON files in the given folder that match the naming pattern "202501241*_PECmd_Output.json",
    allows the user to select one, and displays its contents in a nicely formatted way.
    
    :param folder_path: The directory to scan for JSON files.
    """
    prefix,suffix = analysis_search

    if not os.path.exists(folder_path):
        logging.error(f"Folder not found: {folder_path}")
        print(f"Error: Folder not found: {folder_path}")
        return

    # Collect matching JSON file paths
    json_files = [
        os.path.join(folder_path, file)
        for file in os.listdir(folder_path)
        if file.startswith(prefix) and file.endswith(suffix)
    ]

    if not json_files:
        logging.warning("No matching JSON files found.")
        print("No matching JSON files found.")
        return

    # Display available JSON files
    print("\nAvailable JSON files:")
    for idx, file_path in enumerate(json_files, start=1):
        print(f"{idx}. {os.path.basename(file_path)}")

    # Get user selection
    try:
        choice = int(input("\nEnter the number of the file you want to view: ")) - 1
        if choice < 0 or choice >= len(json_files):
            print("Invalid selection.")
            return
    except ValueError:
        print("Invalid input. Please enter a number.")
        return

    selected_file = json_files[choice]

    # Load and display JSON content
    try:
        with open(selected_file, "r", encoding="utf-8") as f:
            data = json.load(f)

        print("\n---- JSON Content ----")
        print(json.dumps(data, indent=4, sort_keys=True))  # Nicely formatted JSON output
        print("\n----------------------")

        logging.info(f"Displayed JSON file: {selected_file}")

    except (json.JSONDecodeError, FileNotFoundError) as e:
        logging.error(f"Error reading JSON file {selected_file}: {e}")
        print(f"Error: Unable to read the file. {e}")






def compare_json_files(analysis_search,folder_path):
    """
    Compares the 'Directories' and 'FilesLoaded' keys between two selected JSON files 
    and displays only the differences if they are very similar.
    
    :param folder_path: The directory to scan for JSON files.
    """
    prefix,suffix = analysis_search

    if not os.path.exists(folder_path):
        logging.error(f"Folder not found: {folder_path}")
        print(f"Error: Folder not found: {folder_path}")
        return

    # Collect matching JSON file paths
    json_files = [
        os.path.join(folder_path, file)
        for file in os.listdir(folder_path)
        if file.startswith(prefix) and file.endswith(suffix)
    ]

    if len(json_files) < 2:
        logging.warning("Not enough JSON files found for comparison.")
        print("At least two JSON files are required for comparison.")
        return

    # Display available JSON files
    print("\nAvailable JSON files:")
    for idx, file_path in enumerate(json_files, start=1):
        print(f"{idx}. {os.path.basename(file_path)}")

    # Get user selection for two files
    try:
        choice1 = int(input("\nEnter the number of the first file: ")) - 1
        choice2 = int(input("Enter the number of the second file: ")) - 1

        if choice1 < 0 or choice1 >= len(json_files) or choice2 < 0 or choice2 >= len(json_files) or choice1 == choice2:
            print("Invalid selection.")
            return
    except ValueError:
        print("Invalid input. Please enter numbers.")
        return

    file1, file2 = json_files[choice1], json_files[choice2]

    # Load JSON files
    try:
        with open(file1, "r", encoding="utf-8") as f1, open(file2, "r", encoding="utf-8") as f2:
            data1, data2 = json.load(f1), json.load(f2)

        # Extract necessary fields
        exe_hash1 = f"{data1.get('ExecutableName', 'Unknown')}-{data1.get('Hash', 'NoHash')}"
        exe_hash2 = f"{data2.get('ExecutableName', 'Unknown')}-{data2.get('Hash', 'NoHash')}"

        directories1 = set(data1.get("Directories", "").split(",")) if "Directories" in data1 else set()
        directories2 = set(data2.get("Directories", "").split(",")) if "Directories" in data2 else set()

        files_loaded1 = set(data1.get("FilesLoaded", "").split(",")) if "FilesLoaded" in data1 else set()
        files_loaded2 = set(data2.get("FilesLoaded", "").split(",")) if "FilesLoaded" in data2 else set()

        # Calculate differences
        dir_diff1 = directories1 - directories2  # Items in file1 but not in file2
        dir_diff2 = directories2 - directories1  # Items in file2 but not in file1

        files_diff1 = files_loaded1 - files_loaded2  # Items in file1 but not in file2
        files_diff2 = files_loaded2 - files_loaded1  # Items in file2 but not in file1

        # Display results if they are similar but have differences
        print("\nComparison Results:")
        print(f"\n{exe_hash1} vs {exe_hash2}")

        if not dir_diff1 and not dir_diff2 and not files_diff1 and not files_diff2:
            print("The files are nearly identical. No significant differences found.")
        else:
            if dir_diff1 or dir_diff2:
                print("\nDifferences in 'Directories':")
                if dir_diff1:
                    pprint(f"Only in {exe_hash1}: {', '.join(dir_diff1)}")
                if dir_diff2:
                    pprint(f"Only in {exe_hash2}: {', '.join(dir_diff2)}")

            if files_diff1 or files_diff2:
                print("\nDifferences in 'FilesLoaded':")
                if files_diff1:
                    pprint(f"Only in {exe_hash1}: {', '.join(files_diff1)}")
                if files_diff2:
                    pprint(f"Only in {exe_hash2}: {', '.join(files_diff2)}")

        logging.info(f"Compared files: {file1} and {file2}")

    except (json.JSONDecodeError, FileNotFoundError, KeyError) as e:
        logging.error(f"Error reading JSON files: {e}")
        print(f"Error: Unable to read the files. {e}")





def delete_all_files(folder_path):
    """
    Deletes all files in the specified folder, leaving subdirectories intact.

    :param folder_path: Path to the folder where all files will be deleted.
    """
    if not os.path.exists(folder_path):
        logging.error(f"Folder not found: {folder_path}")
        print(f"Error: Folder not found: {folder_path}")
        return

    file_count = 0

    for file in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file)
        
        if os.path.isfile(file_path):  # Ensure it's a file, not a directory
            try:
                os.remove(file_path)
                logging.info(f"Deleted: {file_path}")
                print(f"Deleted: {file_path}")
                file_count += 1
            except Exception as e:
                logging.error(f"Error deleting {file_path}: {e}")
                print(f"Error deleting {file_path}: {e}")

    if file_count == 0:
        print("No files found to delete.")
        logging.warning("No files found to delete.")




def move_specific_files(analysis_search, folder_path, counter):
    """
    Moves all files with specified extensions from source to destination folder.

    :param source_folder: Folder where files will be moved from.
    :param destination_folder: Folder where files will be moved to.
    :param extensions: Tuple of file extensions to move (default: .pf, .json).
    """
    prefix, suffix = analysis_search

    if not os.path.exists(folder_path):
        logging.error(f"Source folder not found: {folder_path}")
        print(f"Error: Source folder not found: {folder_path}")
        return

    if not os.path.exists(folder_path+f"//{counter}"):
        os.makedirs(folder_path+f"//{counter}")  # Create destination if it doesn't exist
        logging.info(f'Created destination folder: {folder_path}+"\\"{counter}')

    file_count = 0

    for file in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file)

        # Check if it's a file and matches the allowed extensions
        if os.path.isfile(file_path) and file.lower().endswith(suffix):
            destination_path = os.path.join(folder_path+f"//{counter}", file)

            try:
                shutil.move(file_path, folder_path+f"//{counter}")
                logging.info(f"Moved: {file_path} -> {folder_path+f"//{counter}"}")
                print(f"Moved: {file} -> {folder_path+f"//{counter}"}")
                file_count += 1
            except Exception as e:
                logging.error(f"Error moving {file_path}: {e}")
                print(f"Error moving {file_path}: {e}")

    if file_count == 0:
        print("No matching files found to move.")
        logging.warning("No matching files found to move.")




def get_menu():
    while True:

        prompt = prompt="""Options: 
                      
                       0 - Copy Prefetch files from default location
                       1 - Only analyze
                       2 - Only extract from json
                       3 - Analyze and extract
                       4 - View json files
                       5 - Compare JSON files (References)
                       8 - SpeedTest (0->1->2->delete) WARN!!
                       9 - SpeedTest x 10 (with looping)
                       X - Exit
                       Enter: """

        user_input = input(prompt)
        print(user_input)
        if user_input is not None:
            return user_input
        else:
            exit





def main():

    PARAMETERS = ""
    #PARAMETERS
    folder = r"C:\\Users\\4n6mole\\Desktop\\Testing\\Prefetch"  # Change this to your destination folder (Work folder)
    cmd_template = r'"C:\\Users\\4n6mole\\Desktop\\Tools\\Get-ZimmermanTools\\net9\\PECmd.exe" -f "{}" --json C:\\Users\\4n6mole\\Desktop\\Testing\\Prefetch --jsonf C:\\Users\\4n6mole\\Desktop\\Testing\\Prefetch\\{}.json'  # Modify with your actual command
    
    #KEYWORDS TO COPY RELEVANT PF FILES
    
    destination = folder 
    source = r"C:\\Windows\\Prefetch"  # Change this to your source folder
    delete_files_in_folder_path = destination #Used for multiple testings to empty e.g. destination path

    #KEYWORDS TO IDENTIFY RELEVANT FILES FOR PROCESSING WITH PRCmd.exe
    pf_copy_prefix = "CHROME.EXE-"
    pf_copy_suffix = ".pf"
    pfcopy_search = (pf_copy_prefix, pf_copy_suffix)

    process_search = (pf_copy_prefix, pf_copy_suffix)

    #KEYWORDS TO IDENTIFY RELEVANT FILES FOR ANALYSIS
    analysis_file_suffix = ".json"
    analysis_search = (pf_copy_prefix, analysis_file_suffix)

    #TEST SUBJECT PROGRAM
    program_path = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"  # Change this path if needed
    program_name = "chrome.exe" #Check in task manager
    counter = 0 #ON WHAT LOOP IS SCRIPT - Can be adjusted to represent actual run count


    #MENU
    menu_input = get_menu() # Call menu - Option 9 uses hidden delete option, it deletes content (.pf and .json files) of folder specified in variable delete_files_in_folder_path
                       
                       
    if menu_input == "X":
        exit
    elif menu_input == "0":
        logging.info("Fetching Prefetch files...")
        copy_prefetch_files(source, destination, pfcopy_search)
        logging.info("Copy operation completed...")
        menu_input = get_menu()

    elif menu_input == "1":
        logging.info("Analyzing Prefetch files...")
        process_files(process_search, folder, cmd_template)
        logging.info("Analysis completed...")
        menu_input = get_menu()

    elif menu_input == "2": 
        logging.info("Extracting data from json files...")
        process_json_files(analysis_search, folder, counter=0)
        logging.info("Extracting from JSON files completed...")
        menu_input = get_menu()

    elif menu_input == "3":
        logging.info("Analyzing Prefetch files...")
        process_files(process_search, folder, cmd_template)
        logging.info("Analysis completed...")
        logging.info("Extracting data from json files...")
        process_json_files(analysis_search, folder, counter=0)
        logging.info("Extracting from JSON files completed...")
        menu_input = get_menu()

    elif menu_input == "4":
        logging.info("Extracting data from json files...")
        view_json_files(analysis_search, folder)
        logging.info("Extraction completed...")
        menu_input = get_menu()
    
    elif menu_input == "5":
        logging.info("Compare JSON files (Directories and Files Loaded)...")
        compare_json_files(analysis_search, folder)
        logging.info("Comparison completed...")
        menu_input = get_menu()

    elif menu_input == "delete":
        logging.info("Compare JSON files (Directories and Files Loaded)...")
        delete_all_files(delete_files_in_folder_path)
        logging.info("Deletion completed...")
        menu_input = get_menu()

    elif menu_input == "8":

        logging.info("Fetching Prefetch files...")
        copy_prefetch_files(source, destination, pfcopy_search)
        logging.info("Copy operation completed...")

        logging.info("Analyzing Prefetch files...")
        process_files(process_search, folder, cmd_template)
        logging.info("Analysis completed...")

        logging.info("Extracting data from json files...")
        process_json_files(analysis_search, folder, counter=0)
        logging.info("Extracting from JSON files completed...")

        logging.info("Compare JSON files (Directories and Files Loaded)...")
        delete_all_files(delete_files_in_folder_path)
        logging.info("Deletion completed...")
        menu_input = get_menu()

    elif menu_input == "9":
        #x5 LOOP - addjust number in line below  
        while counter <= 5: 
            logging.warning(f"### RUN {counter} ###")

            logging.info("Start/Close program...")
            start_and_close_program(program_path,program_name)
            logging.info("Start/Close completed...")

            logging.info("Fetching Prefetch files...")
            copy_prefetch_files(source, destination, pfcopy_search)
            logging.info("Copy operation completed...")

            logging.info("Analyzing Prefetch files...")
            process_files(process_search, folder, cmd_template)
            logging.info("Analysis completed...")

            logging.info("Extracting data from json files...")
            process_json_files(analysis_search, folder,counter)
            logging.info("Extracting from JSON files completed...")
            
            logging.info("Move Prefetch files...")
            move_specific_files(analysis_search, destination, counter)
            logging.info("Move completed...")

            logging.info("Compare JSON files (Directories and Files Loaded)...")
            delete_all_files(delete_files_in_folder_path)
            logging.info("Deletion completed...")
            time.sleep(10)

            counter=counter+1

if __name__ == "__main__":

    main()
