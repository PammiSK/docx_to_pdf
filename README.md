# DOCX to PDF Converter

A simple Python script that converts all `.docx` files in the current directory to `.pdf` format, automatically overwriting existing `.pdf` files.

## Features

- Converts all `.docx` files in the current folder to `.pdf`.
- Overwrites existing `.pdf` files if they already exist.
- Runs automatically without user input.

## Prerequisites

Ensure you have the following installed on your system:

- **Python** (3.x recommended)
- **Microsoft Word** (Required for conversion)
- **Required Python libraries:**

  Install dependencies using pip:
  ```sh
  pip install comtypes
  ```

## Usage

1. **Download or Clone the Repository:**

2. **Run the script:**

   ```sh
   python docx_to_pdf.py
   ```

   - This will automatically convert all `.docx` files in the current directory to `.pdf`.
   - Converted files will have the same name but with a `.pdf` extension.

## Convert to .exe

You can convert this script into a Windows executable for easier use:

```sh
pyinstaller --onefile --noconsole docx_to_pdf.py
```

After running the above command, the executable will be in the `dist/` folder.

## Example Output

```
Converted: document1.pdf
Converted: report.pdf
Converted: notes.pdf
```

## License

This project is licensed under the MIT License.

