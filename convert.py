import os
import win32com.client
import shutil

def convert_to_pdf(input_path, output_path):
    """
    Converts a PowerPoint file to PDF.
    Args:
        input_path (str): Path to the input PowerPoint file.
        output_path (str): Path to save the output PDF.
    """
    # Ensure absolute paths
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    # print(f"Processing file: {input_path}")
    if not os.path.exists(input_path):
        print(f"File not found: {input_path}")

    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1
        presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
        presentation.SaveAs(output_path, 32)  # 32 corresponds to the PDF format
        presentation.Close()
    except Exception as e:
        print(f"Failed to convert {input_path} to PDF: {e}")
    finally:
        powerpoint.Quit()


def process_folder(input_folder, output_folder):
    """
    Recursively processes a folder, converts PowerPoint files to PDF, 
    and maintains the folder structure in the output folder.
    Args:
        input_folder (str): Path to the folder containing PowerPoint files.
        output_folder (str): Path to save converted PDF files.
    """
    for root, _, files in os.walk(input_folder):
        # Create the corresponding output directory
        relative_path = os.path.relpath(root, input_folder)
        target_dir = os.path.join(output_folder, relative_path)
        os.makedirs(target_dir, exist_ok=True)
        
        for file in files:
            if file.lower().endswith(('.ppt', '.pptx', '.pps', '.ppsx')):
                input_file_path = os.path.join(root, file)
                output_file_path = os.path.join(target_dir, os.path.splitext(file)[0] + '.pdf')
                print(f"Converting: {input_file_path} -> {output_file_path}")
                convert_to_pdf(input_file_path, output_file_path)


if __name__ == "__main__":
    import argparse
    
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description="Recursively convert PowerPoint files to PDFs.")
    parser.add_argument("input_folder", help="Path to the folder containing PowerPoint files.")
    parser.add_argument("output_folder", help="Path to save the converted PDF files.")
    args = parser.parse_args()

    # Ensure the output folder is clean
    if os.path.exists(args.output_folder):
        shutil.rmtree(args.output_folder)
    os.makedirs(args.output_folder, exist_ok=True)

    # Process the folder
    process_folder(args.input_folder, args.output_folder)

    print("Conversion completed!")
