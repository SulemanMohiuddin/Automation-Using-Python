import os


def rename_files_in_sequence(folder_path):
    try:
        # Get a list of files in the folder
        files = os.listdir(folder_path)

        # Sort files to ensure consistent ordering
        files.sort()

        # Rename each file in sequence
        for idx, file_name in enumerate(files):
            # Get full file path
            old_file_path = os.path.join(folder_path, file_name)

            # Skip directories
            if os.path.isdir(old_file_path):
                continue

            # Extract file extension
            file_extension = os.path.splitext(file_name)[1]

            # Create new file name with sequence number and '_Lf'
            new_file_name = f"{idx + 1}_LS{file_extension}"

            # Full new file path
            new_file_path = os.path.join(folder_path, new_file_name)

            # Rename the file
            os.rename(old_file_path, new_file_path)

        print(f"Files in {folder_path} have been renamed in the format 'X_Lf'.")

    except Exception as e:
        print(f"An error occurred: {e}")


# Take input for the folder path
folder_path = r"C:\Users\Suleman\Desktop\Lightmode Smallscreen"
rename_files_in_sequence(folder_path)
