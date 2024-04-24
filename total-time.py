import os
import re
import zipfile
import re
import tempfile
from time import sleep

def replace_total_time(content, new_time):
    # Use regular expressions to find and replace the total time
    pattern = r'<TotalTime>(.*?)</TotalTime>'
    replaced_content = re.sub(pattern, f'<TotalTime>{new_time}</TotalTime>', content, flags=re.DOTALL)
    return replaced_content

def edit_xml_in_docx(docx_dir, new_time):
    # Create a temporary directory
    with tempfile.TemporaryDirectory() as tmp_dir:
        # Extract the .docx file to the temporary directory
        with zipfile.ZipFile(docx_dir, 'r') as zip_ref:
            zip_ref.extractall(tmp_dir)
        
        # Open and edit the XML file
        xml_file_path = os.path.join(tmp_dir, 'docProps', 'app.xml')
        with open(xml_file_path, 'r', encoding='utf-8') as file:
            content = file.read()

        # Replace total time with user's chosen time
        replaced_content = replace_total_time(content, new_time)

        # Write the modified content back to the XML file
        with open(xml_file_path, 'w', encoding='utf-8') as file:
            file.write(replaced_content)

        # Re-zip the directory back to .docx format
        with zipfile.ZipFile(docx_dir, 'w') as zip_ref:
            for folder_name, subfolders, filenames in os.walk(tmp_dir):
                for filename in filenames:
                    file_path = os.path.join(folder_name, filename)
                    arcname = os.path.relpath(file_path, tmp_dir)
                    zip_ref.write(file_path, arcname=arcname)

def edit_total_time_in_docx(docx_dir, new_time):
    try:
        edit_xml_in_docx(docx_dir, new_time)
    except Exception as e:
        print(f"An error occurred: {str(e)}")

# Example usage
print("___________     __         ._____________.__                ")
print("\__    ___/____/  |______  |  \__    ___/|__| _____   ____  ")
print("  |    | /  _ \   __\__  \ |  | |    |   |  |/     \_/ __ \ ")
print("  |    |(  <_> )  |  / __ \|  |_|    |   |  |  Y Y  \  ___/ ")
print("  |____| \____/|__| (____  /____/____|   |__|__|_|  /\___  >")
print("                         \/                       \/     \/ \n\n")
print("                                  - by Zanci19 (hvala Petek)")
sleep(2)
os.system('cls' if os.name == 'nt' else 'clear')

docx_dir = input("Enter path to the Word document whose time you want to edit (PATH MUST NOT CONTAIN QUOTATION MARKS): ")
new_time = int(input("Enter new time (in minutes): "))
if os.path.exists(docx_dir):
    warning = input(f"\nYou are about to change the TotalTime property of the Word document. Please make sure you fully understand what you are doing.\nThe Word document resides in the following directory:\n\n{docx_dir}\n\nDocument contains the <TotalTime></TotalTime> property, which will, upon execution of the process, hold the value of exactly {new_time} minutes.\nAre you sure you want to continue?\n[y = yes, n = no]\n").lower()
    os.system('cls' if os.name == 'nt' else 'clear')
else:
    print("Error: The provided path does not exist. Enter a valid path that exists in this system.\nThe program will close in three seconds.")
    sleep(3)

if warning == "y":
    try:
        sleep(1)
        print("The execution will now commence. Do not close the program.")
        sleep(2)
        os.system('cls' if os.name == 'nt' else 'clear')
        edit_total_time_in_docx(docx_dir, new_time)
        print(f"Property <TotalTime></TotalTime> has been updated to {new_time} minutes in the document that exists in the next path:\n\n{docx_dir}\n\nThe execution has completed successfully.")
    except FileNotFoundError:
        print("Error: The provided path does not exist. Enter a valid path that exists in this system.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    print("The program will close in five seconds.")
    sleep(5)
elif warning == "n":
    print("The process was aborted. The execution will not commence.\nThe program will close in three seconds.")
    sleep(3)
else:
    print("Incorrect input has been provided. The process was aborted. The execution will not commence.\nThe program will close in three seconds.")
    sleep(3)
