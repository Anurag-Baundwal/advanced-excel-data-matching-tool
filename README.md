# Excel Investigator Matcher

This Python script showcases advanced Excel and pandas skills by processing an Excel file containing investigator and fire site data, and determining whether the investigators within each cluster (grouped by `fire_investigator_id`) are the same person based or not on certain matching criteria. It adds two new columns to the output Excel file: `same` (indicating if the investigators are the same person) and `matching_criteria` (listing the attributes that match).

## Features

- Matches investigators within each cluster based on name, phone number, email, and location (country, state, city)
- Handles different phone number formats and multiple phone numbers/emails per investigator
- Applies exact matching for names, phone numbers, emails, and location
- Performs matching in multiple passes to identify linked matches and cross-link matches
- Outputs the results to a new Excel file with color-coded rows based on the matching pass

## Matching Process

The script performs the following passes to match investigators within each cluster:

1. **First Pass**: Direct match on phone numbers or emails
2. **Second Pass**: Linked match on phone numbers or emails
3. **Third Pass**: Cross-link matching on phone numbers or emails
4. **Fourth Pass**: Exact match on names and location (country, state, city)

The script reports the number of matches found in each pass and the overall match percentage.

## Requirements

- Python 3.6 or higher
- pandas
- regex
- openpyxl
- fuzzywuzzy
- python-Levenshtein

## Setup

1. Clone the repository:

```bash
git clone https://github.com/Anurag-Baundwal/excel-data-matching-tool
```

2. Navigate to the project directory:

```bash
cd excel-data-matching-tool
```

3. Create a virtual environment:

```bash
python -m venv venv
```

4. Activate the virtual environment:

- On Windows:
  ```bash
  venv\Scripts\activate
  ```
- On macOS and Linux:
  ```bash
  source venv/bin/activate
  ```

5. Install the required packages:

```bash
pip install -r requirements.txt
```

## Usage

1. Place your input Excel file (e.g., `input.xlsx`) in the project directory.

2. Run the script:

```bash
python app.py
```

3. The output Excel file (e.g., `output.xlsx`) will be generated in the project directory, containing the original data along with the two new columns: `same` and `matching_criteria`.

## License

This project is open-source and available under the [MIT License](LICENSE).