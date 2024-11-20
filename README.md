# PowerPoint to PDF Converter

This is a Python script that converts PowerPoint presentations (.ppt, .pptx, .pps, .ppsx) to PDF format. The script recursively processes a folder, maintains the directory structure, and saves the converted PDFs in a new specified output folder. It is designed specifically for use on **Windows** with **Microsoft PowerPoint** installed.

## Features
- Recursively scans input folders and subfolders for PowerPoint files.
- Converts `.ppt`, `.pptx`, `.pps`, and `.ppsx` files to PDF format.
- Maintains the original folder structure in the output folder.
- Provides clear logging during the conversion process.

## Requirements
- **Python 3.x**
- **Microsoft PowerPoint** (installed and activated)
- **`pywin32`** library (used to interact with PowerPoint's COM interface)

## Installation

1. **Clone the Repository**
   ```bash
   git clone https://github.com/pgainullin/ppt-to-pdf-converter.git
   cd ppt-to-pdf-converter
   ```

2. **Install Dependencies**
   Install the required Python package using pip:
   ```bash
   python -m pip install pywin32
   ```

## Usage

To use the script, run it from the command line with the following arguments:

```bash
python convert_to_pdf.py <input_folder> <output_folder>
```

- `<input_folder>`: Path to the folder containing PowerPoint files you want to convert.
- `<output_folder>`: Path to save the converted PDF files. The folder structure will be maintained.

### Example

```bash
python convert_to_pdf.py "C:\path\to\input_folder" "C:\path\to\output_folder"
```
This will:
- Convert all PowerPoint files located in `C:\path\to\input_folder`.
- Save the PDFs to `C:\path\to\output_folder`, preserving the original subfolder structure.

## Notes
- The script uses **absolute paths** to prevent path-related errors.
- Ensure you have appropriate permissions for reading files from the input folder and writing files to the output folder.
- Make sure PowerPoint files are not locked or open by another process during the conversion.

## Troubleshooting
- **ModuleNotFoundError: No module named 'win32com'**: Ensure `pywin32` is installed. Run `python -m pip install pywin32`.
- **PowerPoint COM Errors**: These errors may be due to invalid file paths, corrupted files, or missing PowerPoint installations. Ensure all file paths are correct and PowerPoint is properly installed and activated.

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing
If you'd like to contribute, feel free to fork the repository and submit a pull request. All contributions are welcome!

## Contact
If you have any questions or need help, please feel free to open an issue or contact me at [pg@palm83.com].