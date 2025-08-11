# QR Code Generator

![Python](https://img.shields.io/badge/Python-3.7%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)

A professional desktop application for generating and customizing QR codes from data in Excel or CSV files. The application provides a user-friendly interface to create QR codes in various formats, tailored to your needs.

## Features

- **User-Friendly GUI**: A clean and intuitive interface built with `tkinter`.
- **Data Import**: Import data directly from Excel (`.xlsx`) or CSV (`.csv`) files.
- **Column Selection**: Easily select the column containing the data for QR code generation.
- **Multiple Export Formats**:
    - **PDF**: Generate a multi-page PDF with a grid of QR codes, perfect for printing.
    - **PNG**: Export individual QR codes as high-quality PNG images.
    - **ZIP**: Create a ZIP archive containing all generated QR codes as PNG images.
- **Advanced Customization**:
    - **Size**: Adjust the size of the QR codes in pixels.
    - **Colors**: Choose custom foreground and background colors.
    - **Logo Integration**: Add your own logo to the center of the QR codes.
- **Generation Modes**:
    - **Text Mode**: For general-purpose QR codes with text-based data.
    - **Numeric Mode**: Specialized mode for numeric data, with options to prepend or append numbers.
- **Responsive Interface**: Asynchronous processing ensures the application remains responsive during QR code generation.
- **Real-Time Progress**: A progress bar and status updates keep you informed during the generation process.

## Requirements

To run this application, you will need Python 3.7 or higher and the following libraries:

- `pandas`
- `qrcode`
- `reportlab`
- `Pillow`
- `openpyxl`

You can install them using pip:
```bash
pip install pandas qrcode reportlab pillow openpyxl
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
7.  **Select the export format** (PDF, PNG, ZIP).
8.  **Click "Gerar QR Codes"** and choose a location to save the generated file(s).

## Screenshots

*(Placeholder for application screenshots)*

## Known Issues

-   **SVG Export**: The option to export QR codes as SVG files is present in the UI, but the functionality is not yet implemented.

## Credits

This application was developed by **Johann Sebastian Dulz**.
