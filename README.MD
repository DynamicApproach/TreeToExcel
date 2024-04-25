# Python Directory Tree Parser (Tree command to Excel)

This script is a Python program that parses a directory tree file and generates an Excel file with the directory structure. The script uses threading to perform parsing and Excel file creation concurrently for efficiency.

## Dependencies

The script requires the following Python libraries:

- `re`
- `openpyxl`
- `queue`
- `tqdm`

## How to Run

1. Generate a directory tree file using the command `tree --charset ascii -flashuR /path/to/directory/ > file.txt` in the command prompt.
   e.g. `tree --charset ascii -flashuR / > tree.txt`
2. Run the script. A file dialog will open for you to select the generated tree file.
3. The script will parse the tree file and generate an Excel file with the directory structure. If the number of rows exceeds the Excel row limit, multiple files will be created.

## Functions

- `parse_tree_file(filename, queue)`: Parses the tree file and puts the parsed data into a queue.
- `create_excel(queue, output_filename)`: Creates an Excel file from the data in the queue.
- `process_file(filename)`: Sets up the threads for parsing and Excel file creation.
- `main()`: Opens a file dialog for the user to select the tree file and starts the processing of the file.

## Note

The script assumes that every 4 spaces or dashes in the tree file denote a new level in the directory structure. If your tree file uses a different number of spaces or dashes, you will need to adjust the script accordingly.
