# Data Extractor

This Python script extracts email addresses and phone numbers from PDF, DOCX, and DOC files in a specified directory and saves the extracted data to an Excel file.

## Installation

1. Clone this repository:

    ```bash
    git clone https://github.com/lumeirne/Data-Extractor.git
    ```

2. Navigate to the project directory:

    ```bash
    cd Data-Extractor
    ```

3. Install the required dependencies:

    ```bash
    pip install -r requirements.txt
    ```

## Usage

1. Place your PDF, DOCX, and DOC files in the `input` directory.

2. Run the script:

    ```bash
    python data_extractor.py
    ```

3. The extracted data will be saved to an Excel file named `Output.xls` in the `output` directory.

## Example

Suppose you have the following files in the `input` directory:

- `resume.pdf`
- `cv.docx`
- `contact.doc`

After running the script, the extracted data will be saved to `Output.xls` in the `output` directory.

## Dependencies

- pandas
- PyPDF2
- python-docx
- openpyxl

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
