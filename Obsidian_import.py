import pandas as pd
import os
from datetime import datetime
import openpyxl
from openpyxl import load_workbook
import logging
import re
import fitz  # PyMuPDF
from PIL import Image
import traceback
from tqdm import tqdm

# Set up logging
logging.basicConfig(filename='import_log.txt', level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Define the template as a string
TEMPLATE = """---
date: {{Date}}
article: "{{Article}}"
theme1: {{T}}
theme2: {{Theme-2}}
theme3: {{Theme-3}}
theme4: {{Theme-4}}
theme5: {{Theme-5}}
name1: {{Name-1}}
name2: {{Name-2}}
place1: {{Place-1}}
place2: {{Place-2}}
source: {{Src}}
format: {{Fmt}}
transcribed: {{Transcribed}}
last_imported: {{last_imported}}
---

## Article Description

{{Article}}

## Source Information

- Newspaper: {{Newspaper-or-Source}}
- Published: {{Published}}

## Links

- [Online Source]({{Web}})
- [[{{Full-Filename}}|Local File]]

## Thumbnail

![[{{thumbnail}}]]

## Transcription

{{transcription}}

## User Comments

{{user_comments}}

## Tags

#Llanychan #Taber-Project #{{Src}} {{additional_tags}}
"""


def clean_filename(filename):
    # Remove any characters that are not alphanumeric, space, comma, or period
    cleaned = re.sub(r'[^\w\s,.-]', '', filename)
    # Replace multiple spaces with a single space
    cleaned = re.sub(r'\s+', ' ', cleaned)
    return cleaned.strip()

def create_thumbnail(file_path, output_path, size=(300, 300)):
    try:
        if file_path.lower().endswith('.pdf'):
            # Handle PDF files
            doc = fitz.open(file_path)
            pix = doc[0].get_pixmap()
            pix.save(output_path)

            # Resize the saved image
            with Image.open(output_path) as img:
                img.thumbnail(size)
                img.save(output_path)
        else:
            # Handle image files
            with Image.open(file_path) as img:
                if img.mode in ('RGBA', 'LA', 'P'):
                    img = img.convert('RGB')
                img.thumbnail(size)
                img.save(output_path, 'JPEG')

        logging.info(f"Thumbnail created: {output_path}")
        return os.path.basename(output_path)
    except Exception as e:
        logging.error(f"Error creating thumbnail for {file_path}: {str(e)}")
        logging.debug(traceback.format_exc())
        return None


def create_note(row, output_dir):
    try:
        logging.info(f"Starting to create note for article: {row['Article']}")

        # Parse the date
        try:
            date_obj = datetime.strptime(row['Date'], '%Y-%m-%d')
        except ValueError as e:
            logging.error(f"Invalid date format for article: {row['Article']}. Date: {row['Date']}")
            return None

        # Remove year from the beginning of the Article
        article_without_year = re.sub(f'^{date_obj.year}\s*', '', row['Article'])

        # Create filename
        filename = f"{row['Date']} {clean_filename(article_without_year)}.md"

        # Ensure the filename isn't too long (adjust max_length as needed)
        max_length = 255  # Maximum filename length for most file systems
        if len(filename) > max_length:
            filename = filename[:max_length - 3] + '.md'

        filepath = os.path.join(output_dir, filename)

        # Populate the template
        content = TEMPLATE
        for key, value in row.items():
            content = content.replace(f"{{{{{key}}}}}", str(value) if pd.notna(value) else "")

        content = content.replace("{{last_imported}}", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        content = content.replace("{{transcription}}", "")
        content = content.replace("{{user_comments}}", "")
        content = content.replace("{{additional_tags}}", "")

        # Handle thumbnail
        thumbnail_created = False
        local_file_path = row['Full-Filename']
        if pd.notna(local_file_path) and os.path.exists(local_file_path):
            thumbnail_filename = f"thumb_{os.path.basename(local_file_path)}.jpg"
            thumbnail_path = os.path.join(output_dir, 'thumbnails', thumbnail_filename)
            os.makedirs(os.path.dirname(thumbnail_path), exist_ok=True)
            thumbnail_result = create_thumbnail(local_file_path, thumbnail_path)
            if thumbnail_result:
                content = content.replace("{{thumbnail}}", thumbnail_result)
                thumbnail_created = True
            else:
                content = content.replace("![[{{thumbnail}}]]", "")
        else:
            content = content.replace("![[{{thumbnail}}]]", "")
            logging.warning(f"No local file found for article: {row['Article']}")

        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)

        logging.info(f"Successfully created note: {filename}")
        return {"success": True, "thumbnail_created": thumbnail_created}
    except Exception as e:
        logging.error(f"Error creating note for {row['Article']}: {str(e)}")
        logging.debug(traceback.format_exc())
        return None


def main():
    input_file = r"G:\Projects\Clwyd Hall\_Resources\Clwyd Hall Project.xlsm"
    output_dir = r"G:\Projects\Obsidian\Vaultez\Newspapers"
    sheet_name = "Newspapers"

    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(os.path.join(output_dir, 'thumbnails'), exist_ok=True)

    try:
        logging.info(f"Reading Excel file: {input_file}")
        df = pd.read_excel(input_file, sheet_name=sheet_name, engine='openpyxl')
        logging.info(f"Successfully read {len(df)} records from Excel")

        # Filter rows where Src is NC, BN, or WN
        df_filtered = df[df['Src'].isin(['NC', 'BN', 'WN'])]

        # Update Full-Filename
        df_filtered['Full-Filename'] = 'G:\\Projects\\_Resources\\Newspapers\\images\\' + df_filtered['Full-Filename']

        notes_created = 0
        thumbnails_created = 0

        for index, row in tqdm(df_filtered.iterrows(), total=len(df_filtered), desc="Processing records"):
            try:
                logging.debug(f"Processing row {index + 1}")
                result = create_note(row, output_dir)
                if result and result['success']:
                    notes_created += 1
                    if result['thumbnail_created']:
                        thumbnails_created += 1
            except Exception as e:
                logging.error(f"Error processing row {index + 1}: {str(e)}")
                logging.debug(traceback.format_exc())

        print(f"\nProcessed {len(df_filtered)} records")
        print(f"Notes created: {notes_created}")
        print(f"Thumbnails created: {thumbnails_created}")
        logging.info(f"Processed {len(df_filtered)} records")
        logging.info(f"Notes created: {notes_created}")
        logging.info(f"Thumbnails created: {thumbnails_created}")
    except Exception as e:
        logging.error(f"Error processing Excel file: {str(e)}")
        logging.debug(traceback.format_exc())


if __name__ == "__main__":
    main()
