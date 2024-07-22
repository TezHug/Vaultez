import pandas as pd
import os
from datetime import datetime
import logging
from tqdm import tqdm

# Set up logging
logging.basicConfig(filename='census_import_log.txt', level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Define the template as a string
TEMPLATE = """---
date: {{Date}}
article: "{{Article}}"
theme1: {{T}}
{{additional_properties}}
---

## Census Information

{{Article}}

## Location

{{Place_1}}

## Links

- [Online Source]({{Web}})
- [[{{Full_Filename}}|Local File]]

## Tags

#Llanychan #Taber-Project #Census #{{Date}}
"""

def clean_filename(filename):
    # Remove any characters that are not alphanumeric, space, or hyphen
    return ''.join(c for c in filename if c.isalnum() or c in [' ', '-']).strip()

def create_note(row, output_dir):
    try:
        logging.info(f"Starting to create note for census: {row['Article']}")

        # Create filename (without date)
        filename = f"{clean_filename(row['Article'])}.md"

        # Ensure the filename isn't too long (adjust max_length as needed)
        max_length = 255  # Maximum filename length for most file systems
        if len(filename) > max_length:
            filename = filename[:max_length - 3] + '.md'

        filepath = os.path.join(output_dir, filename)

        # Populate the template
        content = TEMPLATE
        additional_properties = ""
        for key, value in row.items():
            if pd.notna(value) and key not in ['Date', 'Article', 'T', 'Src', 'Fmt']:
                additional_properties += f"{key}: {value}\n"
            content = content.replace(f"{{{{{key}}}}}", str(value) if pd.notna(value) else "")

        content = content.replace("{{additional_properties}}", additional_properties.strip())

        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)

        logging.info(f"Successfully created note: {filename}")
        return True

    except Exception as e:
        logging.error(f"Error creating note for {row['Article']}: {str(e)}")
        return False

def main():
    input_file = r"G:\Projects\Clwyd Hall\_Resources\Clwyd Hall Project.xlsm"
    output_dir = r"G:\Projects\Obsidian\Vaultest\Census"
    sheet_name = "Newspapers"

    os.makedirs(output_dir, exist_ok=True)

    try:
        logging.info(f"Reading Excel file: {input_file}")
        df = pd.read_excel(input_file, sheet_name=sheet_name, engine='openpyxl')
        logging.info(f"Successfully read {len(df)} records from Excel")

        # Filter rows where Src is CEN
        df_filtered = df[df['Src'] == 'CEN']

        notes_created = 0

        for index, row in tqdm(df_filtered.iterrows(), total=len(df_filtered), desc="Processing census records"):
            try:
                logging.debug(f"Processing row {index + 1}")
                if create_note(row, output_dir):
                    notes_created += 1
            except Exception as e:
                logging.error(f"Error processing row {index + 1}: {str(e)}")

        print(f"\nProcessed {len(df_filtered)} census records")
        print(f"Notes created: {notes_created}")

        logging.info(f"Processed {len(df_filtered)} census records")
        logging.info(f"Notes created: {notes_created}")

    except Exception as e:
        logging.error(f"Error processing Excel file: {str(e)}")

if __name__ == "__main__":
    main()
