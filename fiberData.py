import os
import subprocess

def main():
    # Ask the user for the folder containing .sor files
    input_folder = input("Enter the path to the folder containing .sor files: ").strip()
    while not os.path.isdir(input_folder):
        print("The provided path does not exist or is not a directory. Please try again.")
        input_folder = input("Enter the path to the folder containing .sor files: ").strip()

    # Ask the user for the folder to save parsed files
    output_folder = input("Enter the path to the folder where parsed files should be saved: ").strip()
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"Created output folder: {output_folder}")

    # Ask the user for the path to the rbOTDR executable
    rbOTDR_path = input("Enter the path to the rbOTDR executable: ").strip()
    while not os.path.isfile(rbOTDR_path) or not os.access(rbOTDR_path, os.X_OK):
        print("The provided rbOTDR path does not exist or is not executable. Please try again.")
        rbOTDR_path = input("Enter the path to the rbOTDR executable: ").strip()

    # Process each .sor file in the input folder
    for root, _, files in os.walk(input_folder):
        for file in files:
            if file.lower().endswith('.sor'):
                input_file_path = os.path.join(root, file)
                output_file_path = os.path.join(output_folder, file)

                print(f"Processing {input_file_path}...")
                try:
                    # Run the rbOTDR command and save output
                    with open(output_file_path, 'w') as output_file:
                        subprocess.run([rbOTDR_path, input_file_path], stdout=output_file, check=True)
                    print(f"Saved parsed data to {output_file_path}")
                except subprocess.CalledProcessError as e:
                    print(f"Error processing {input_file_path}: {e}")
                except Exception as e:
                    print(f"Unexpected error: {e}")

    print("All files processed.")

if __name__ == "__main__":
    main()
