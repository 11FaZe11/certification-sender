# PDF Certification Sender

A Python-based tool to automatically generate personalized PDF certificates from an Excel spreadsheet. This script overlays text (like names) onto a template PDF, allowing for batch creation of custom documents.

## Features

- **Batch Processing**: Generate hundreds of personalized PDFs from a single Excel file.
- **Custom Placement**: Define a precise bounding box (x1, y1, x2, y2) to place text anywhere on the PDF.
- **Text Wrapping**: Long names or text automatically wrap to fit within the defined bounding box.
- **Full Font Support**: Utilizes all TrueType fonts installed on your system. A `FONTS_README.md` is automatically generated to list all available fonts.
- **Custom Styling**: Choose any font, font size, and color to match your certificate's design.
- **User-Friendly Interface**: An interactive command-line interface guides you through the process.
- **Smart Defaults**: Professional default filenames (`input_data.xlsx`, `certificate_template.pdf`) make setup intuitive.
- **Automatic Capitalization**: Names from the Excel file are automatically formatted to title case (e.g., "john doe" becomes "John Doe").

## Prerequisites

- Python 3.6+

## Installation

1.  **Clone the repository:**
    ```bash
    git clone <your-repository-url>
    cd certification-sender
    ```

2.  **Install the required packages:**
    ```bash
    pip install -r requirements.txt
    ```

## How to Use

1.  **Prepare your data:**
    -   Rename your Excel file to `input_data.xlsx` or enter its name when prompted.
    -   Ensure your Excel file has at least two columns: one for the names to be added to the certificate and one for a unique identifier (like a phone number), which will be used for the output filename.

2.  **Prepare your template:**
    -   Rename your PDF certificate template to `certificate_template.pdf` or enter its name when prompted.

3.  **Run the script:**
    ```bash
    python main.py
    ```

4.  **Follow the on-screen prompts:**
    -   **Excel file name**: Press Enter to use the default (`input_data.xlsx`) or type a new name.
    -   **PDF template file name**: Press Enter to use the default (`certificate_template.pdf`) or type a new name.
    -   **Output folder name**: Press Enter to use the default (`output`) or type a new name.
    -   **Column for name**: Enter the Excel column letter for the names (e.g., `A`).
    -   **Column for phone number**: Enter the Excel column letter for the unique identifier (e.g., `C`).
    -   **Bounding Box Coordinates**:
        -   Enter the `x1`, `y1`, `x2`, and `y2` coordinates to define the area where the text will be placed.
        -   The origin `(0,0)` is at the **bottom-left** corner of the PDF page. `x1, y1` is the bottom-left corner of your box, and `x2, y2` is the top-right.
    -   **Font Name**: The script will generate a `FONTS_README.md` file listing all available fonts on your system. Copy a font name from this file and paste it into the terminal.
    -   **Font Size**: Enter the desired font size in points (e.g., `24`).

5.  **Find your certificates:**
    -   The generated PDFs will be saved in the `output` folder (or the custom folder you specified), with filenames corresponding to the unique identifiers from your Excel sheet.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
