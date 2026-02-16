# QR Code Generator

![Python](https://img.shields.io/badge/Python-3.7%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)

A professional desktop application for generating and customizing QR codes **and Code128 barcodes** from data in Excel or CSV files. The application provides a user-friendly interface to create visual codes in various formats, tailored to your needs.

## Features

- **User-Friendly GUI**: A clean and intuitive interface built with `tkinter`.
- **Data Import**: Import data directly from Excel (`.xlsx`) or CSV (`.csv`) files.
- **Column Selection**: Easily select the column containing the data for QR code generation.
- **Multiple Export Formats**:
    - **PDF**: Generate a multi-page PDF with a grid of QR codes, perfect for printing.
    - **PNG**: Export individual QR codes as high-quality PNG images.
    - **ZIP**: Create a ZIP archive containing all generated QR codes as PNG images.
- **Advanced Customization**:
    - **Size**: Adjust QR/barcode width and height in centimeters, with optional "keep ratio" toggles.
    - **Colors**: Choose custom foreground and background colors.
    - **Logo Integration**: Add your own logo to the center of the QR codes.
- **Generation Modes**:
    - **Text Mode**: For general-purpose codes with text-based data.
    - **Numeric Mode**: Specialized mode for numeric data, with options to prepend or append numbers.
- **Code Type Selection**:
    - **QR Code**: Traditional QR generation.
    - **Barcode (Code128)**: Linear barcode option for labels and inventory.
- **Responsive Interface**: Asynchronous processing ensures the application remains responsive during QR code generation.
- **Real-Time Progress**: A progress bar and status updates keep you informed during the generation process.
- **Document Preview (1 page)**: Preview area can render a one-page layout based on the selected column before export.
- **Structured Logging**: Operational events and errors are recorded in `logs/app.log` (JSON lines) for support/diagnostics.

## Requirements

To run this application, you will need Python 3.7 or higher and the following libraries:

- `pandas`
- `qrcode`
- `reportlab`
- `Pillow`
- `openpyxl`
- `python-barcode` *(optional, recommended for Code128 without renderPM backend)*

You can install them using pip:
```bash
pip install pandas qrcode reportlab pillow openpyxl python-barcode
```

## How to Use

1.  **Clone or download the repository.**
2.  **Install the required libraries** (see the "Requirements" section).
3.  **Run the application:**
    ```bash
    python qr_generator.py
    ```
4.  **Select your data file** (Excel or CSV) using the "Selecionar Arquivo" button.
5.  **Choose the column** that contains the data for the QR codes.
6.  **Customize the QR code settings** as needed (size, color, logo, etc.).
7.  **Select the export format** (PDF, PNG, ZIP, SVG for QR) using the visible "Formato de saída" selector.
8.  **Click "Gerar QR Codes"** and choose a location to save the generated file(s).

## Screenshots

*(Placeholder for application screenshots)*

## Known Issues

-   **Barcode SVG Export**: SVG export is currently available for QR Code only.

## Credits

This application was developed by **Johann Sebastian Dulz**.
