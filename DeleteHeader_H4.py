import os

def delete_lines_and_first_tab(file_path, lines_to_delete):
    # Read the file and store the lines in a list
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    
    # Delete the specified lines
    lines = [line for i, line in enumerate(lines) if i not in lines_to_delete]
    
    # Remove the first tab character from each remaining line
    modified_lines = [line.replace('\t', '', 1) if '\t' in line else line for line in lines]
    
    return modified_lines

def process_files(input_directory, output_directory, lines_to_delete):
    # Ensure the output directory exists
    os.makedirs(output_directory, exist_ok=True)
    
    # Process each file in the input directory
    for filename in os.listdir(input_directory):
        if filename.endswith(".txt"):
            input_file_path = os.path.join(input_directory, filename)
            output_file_path = os.path.join(output_directory, filename)
            
            # Delete the specified lines and the first tab character from each line
            modified_lines = delete_lines_and_first_tab(input_file_path, lines_to_delete)
            
            # Write the modified content to the new file in the output directory
            with open(output_file_path, 'w', encoding='utf-8') as file:
                file.writelines(modified_lines)
            
            print(f"Processed {filename}")

if __name__ == "__main__":
    # Specify the input directory containing the text files
    input_directory = r"C:\Users\wasurat.boonnan\OneDrive - Accenture\Desktop\Working Space"
    
    # Specify the output directory where the modified files will be saved
    output_directory = r"C:\Users\wasurat.boonnan\OneDrive - Accenture\Desktop\Working Space\delete header"
    # output_directory = r"C:\Users\wasurat.boonnan\OneDrive - Accenture\Desktop\Data Trnasfomation\Transformed Data\Mock3\M3 Reconcile\M3 Reconcile Export file"

    # Specify the lines to delete (0-based index)
    lines_to_delete = [0, 2]

    
    process_files(input_directory, output_directory, lines_to_delete)