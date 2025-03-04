# QA-to-XLSX Converter

A simple Bash script that converts text files containing question-answer pairs into Excel (XLSX) files.

## Features

- Extracts Q&A pairs from text files (format: "Q: question" followed by "A: answer")
- Creates an Excel file with each Q&A pair in a single cell
- Each Q&A pair is placed in a new row
- Automatically adjusts column width and row height for better readability
- Automatically installs required dependencies
- Outputs the Excel file in the same directory as the input file

## Requirements

- Bash shell
- Python 3
- Internet connection (for first-time dependency installation)

The script will automatically install the required Python packages:
- openpyxl (Excel file handling library)

## Installation

1. Download the script

2. Make it executable:

```bash
chmod +x qa_to_xlsx.sh
```

## Usage

```bash
./qa_to_xlsx.sh <input_text_file>
```

Example:

```bash
./qa_to_xlsx.sh interview_data.txt
```

This will create `interview_data.xlsx` in the same directory as your input file.

## Input Format

The script expects the input text file to have Q&A pairs in the following format:

```
Q: What is your favorite programming language?
A: I really enjoy working with Python because of its simplicity and extensive libraries.

Q: How did you learn to code?
A: I started with online tutorials and then took some computer science courses in college.
```

Each Q&A pair should:
- Start with "Q:" for questions
- Have "A:" for answers
- Pairs are typically separated by an empty line

## Output Format

The script creates an Excel file (.xlsx) in the same directory as your input file. Here's how the output is structured:

```
| Q&A Pair (Column A)                                       |
|-----------------------------------------------------------|
| Q: What is your favorite programming language?            |
| A: I really enjoy working with Python because of its      |
|    simplicity and extensive libraries.                    |
|-----------------------------------------------------------|
| Q: How did you learn to code?                             |
| A: I started with online tutorials and then took some     |
|    computer science courses in college.                   |
|-----------------------------------------------------------|
| Q: What project are you most proud of?                    |
| A: The inventory management system I built that saved     |
|    our company significant time in tracking products.     |
|-----------------------------------------------------------|
```

Key characteristics of the output Excel file:
- Single column format (Column A)
- Row 1: Header row with the title "Q&A Pair"
- Rows 2 and onwards: Each row contains one complete Q&A pair in a single cell
- All text formatting is preserved, including the "Q:" and "A:" prefixes
- Text wrapping is enabled within cells for better readability
- Column width and row heights are automatically adjusted based on content

For example, if your input file contains 5 Q&A pairs, the resulting Excel file will have:
- 1 header row
- 5 data rows (one per Q&A pair)
- Each Q&A pair is contained entirely within a single cell

## Troubleshooting

If you encounter permission issues, try running the script with sudo:

```bash
sudo ./qa_to_xlsx.sh <input_text_file>
```

If the script fails to detect Q&A pairs, ensure your text file follows the expected format with "Q:" and "A:" prefixes.

## License

MIT

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.
