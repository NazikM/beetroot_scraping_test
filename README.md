# Publication Extractor

The **Publication Extractor** is a Python script designed to extract information from PDF documents related to publications, particularly Abstract Book from 5th World..... It utilizes the **pdfplumber** library for PDF text extraction and data manipulation, and the **ExcelWriter** module for exporting the extracted data to an Excel file.

## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [Usage](#usage)
- [Dependencies](#dependencies)
- [Installation](#installation)
- [Configuration](#configuration)
- [License](#license)


## Features

- Extracts publication information from specified PDF files with same structure.
- Processes pages in the PDF document to extract relevant data.
- Handles different font sizes and styles to correctly identify and categorize information.
- Creates an Excel file with the extracted data for easy storage and analysis.

## Usage

1. Place the PDF file containing the conference abstracts in the same directory as the script.
2. Modify the `file_name` variable in the script to match the name of your PDF file.
3. Run the script.
4. The extracted publication information will be stored in an Excel file named `result.xlsx` in the same directory.

## Dependencies

- [pdfplumber](https://github.com/jsvine/pdfplumber): A library for extracting text and metadata from PDF files.
- ExcelWriter (Custom module): A module for writing data to Excel files.

## Installation

To get started with the Beetroot Scraping Test project, follow these steps to set up the required environment and run the `main.py` file on different systems. This guide assumes you have Python and Git already installed on your machine.

### Clone the Repository

Open a terminal/command prompt and navigate to the directory where you want to store the project. Then, run the following command to clone the repository:

```bash
git clone https://github.com/NazikM/beetroot_scraping_test.git
```

### Create a Virtual Environment (Optional but Recommended)

It's a good practice to use a virtual environment to isolate project dependencies. Navigate into the project directory and create a virtual environment:

```bash
cd beetroot_scraping_test
python -m venv venv
```

Activate the virtual environment:

- On Windows (Command Prompt):
  ```bash
  venv\Scripts\activate
  ```

- On macOS and Linux:
  ```bash
  source venv/bin/activate
  ```

### Install Requirements

While in the project directory and with your virtual environment active, install the required packages using `pip`:

```bash
pip install -r requirements.txt
```

### Run the Script

Now that the environment is set up and the dependencies are installed, you can run the `main.py` script:

```bash
python main.py
```


### Configuration

No additional configuration is required if using the provided `ExcelWriter` module. However, if you need to customize the output format or file naming, you may need to modify the `save_to_excel()` function or the `ExcelWriter` module.

## License

This script is provided under the [MIT License](LICENSE).
