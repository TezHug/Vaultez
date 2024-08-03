import pandas as pd
import os
from datetime import datetime
import openpyxl
import logging
import re
import fitz  # PyMuPDF
from PIL import Image
import traceback
from tqdm import tqdm

# Set up logging
logging.basicConfig(filename='newspaper_import_log.txt', level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Define the template for markdown notes
TEMPLATE = """---
cssclass: collapse-properties
date: {{Date}}
article: "{{Article}}"
theme1: {{T}}
theme2: {{Theme_2}}
theme3: {{Theme_3}}
theme4: {{Theme_4}}
theme5: {{Theme_5}}
people:
  name1: {{Name_1}}
  name2: {{Name_2}}
places:
  place1: {{Place_1}}
  place2: {{Place_2}}
source: {{Src}}
format: {{Fmt}}
transcribed: {{Transcribed}}
last_imported: {{last_imported}}
---

# Article
{{Article}}

## People Involved
{{people_involved}}

## Locations
{{locations}}

## Thumbnail
![[{{thumbnail}}]]

## Source Information
- Newspaper: {{Newspaper_or_Source}}
- Published: {{Published}}
- [Online Source]({{Address}})

## Links
- [[{{Full_Filename}}|Local File]]

## Tags
{{tags}}
"""


def clean_filename(filename):
    # Remove or replace invalid characters
    invalid_chars = r'[<>:"/\\|?*]'
    cleaned = re.sub(invalid_chars, '', filename)
    # Replace multiple spaces with a single space
    cleaned = re.sub(r'\s+', ' ', cleaned)
    # Remove leading and trailing spaces
    cleaned = cleaned.strip()
    # Ensure the filename is not empty and doesn't start with a dot
    if not cleaned or cleaned.startswith('.'):
        cleaned = 'untitled_file'
    return cleaned


def create_thumbnail(file_path, thumbnails_dir, size=(300, 300)):
    try:
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        thumbnail_filename = f"thumb_{clean_filename(base_name)}.jpg"
        thumbnail_path = os.path.join(thumbnails_dir, thumbnail_filename)

        if file_path.lower().endswith('.pdf'):
            doc = fitz.open(file_path)
            pix = doc[0].get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img.thumbnail(size)
            img.save(thumbnail_path, 'JPEG')
        else:
            with Image.open(file_path) as img:
                if img.mode in ('P', 'RGBA', 'LA'):
                    img = img.convert('RGB')
                img.thumbnail(size)
                img.save(thumbnail_path, 'JPEG')

        logging.info(f"Thumbnail created: {thumbnail_path}")
        return thumbnail_filename
    except Exception as e:
        logging.error(f"Error creating thumbnail for {file_path}: {str(e)}")
        logging.debug(traceback.format_exc())
        return None


def create_note(row, articles_dir, thumbnails_dir, images_dir):
    try:
        logging.info(f"Starting to create note for article: {row['Article']}")

        # Create filename
        filename = f"{clean_filename(row['Article'])}.md"
        if len(filename) > 255:
            filename = filename[:252] + '.md'
        filepath = os.path.join(articles_dir, filename)

        # Populate the template
        content = TEMPLATE
        for key, value in row.items():
            if pd.notna(value):
                content = content.replace(f"{{{{{key}}}}}", str(value))
            else:
                content = content.replace(f"{{{{{key}}}}}", "")

        content = content.replace("{{last_imported}}", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

        # Process people involved
        people_involved = []
        if pd.notna(row.get('Name_1')):
            people_involved.append(f"- {row['Name_1']}")
        if pd.notna(row.get('Name_2')):
            people_involved.append(f"- {row['Name_2']}")
        content = content.replace("{{people_involved}}", "\n".join(people_involved))

        # Process locations
        locations = []
        if pd.notna(row.get('Place_1')):
            locations.append(f"- {row['Place_1']}")
        if pd.notna(row.get('Place_2')):
            locations.append(f"- {row['Place_2']}")
        content = content.replace("{{locations}}", "\n".join(locations))

        # Process tags
        tags = [f"#Source-{row.get('Src', '')}"]

        # Add Theme tags
        if pd.notna(row.get('T')):
            tags.append(f"#Theme-{row['T'].replace(' ', '-')}")
        for i in range(2, 6):  # This will cover Theme_2 to Theme_5
            if pd.notna(row.get(f'Theme_{i}')):
                tags.append(f"#Theme-{row[f'Theme_{i}'].replace(' ', '-')}")

        # Add Person tags
        for i in range(1, 3):  # This will cover Name_1 and Name_2
            if pd.notna(row.get(f'Name_{i}')):
                tags.append(f"#Person-{row[f'Name_{i}'].replace(' ', '-')}")

        # Add Place tags
        for i in range(1, 3):  # This will cover Place_1 and Place_2
            if pd.notna(row.get(f'Place_{i}')):
                tags.append(f"#Place-{row[f'Place_{i}'].replace(' ', '-')}")

        # Add Year tag
        if pd.notna(row.get('Date')):
            tags.append(f"#Year-{row['Date'][:4]}")  # Assuming Date is in YYYY-MM-DD format

        content = content.replace("{{tags}}", " ".join(tags))

        # Handle thumbnail and local file link
        thumbnail_created = False
        local_file_path = os.path.join(images_dir, row.get('Full_Filename', ''))
        logging.debug(f"Checking for file: {local_file_path}")
        if pd.notna(local_file_path) and os.path.exists(local_file_path):
            file_name = clean_filename(os.path.basename(local_file_path))
            clean_local_path = f"Images/{file_name}"
            content = content.replace("[[{{Full_Filename}}|Local File]]", f"[[{clean_local_path}|Local File]]")

            thumbnail_result = create_thumbnail(local_file_path, thumbnails_dir)
            if thumbnail_result:
                clean_thumbnail_name = clean_filename(thumbnail_result)
                content = content.replace("{{thumbnail}}", f"thumbnails/{clean_thumbnail_name}")
                thumbnail_created = True
            else:
                content = content.replace("![[{{thumbnail}}]]", "")
                logging.warning(f"Failed to create thumbnail for: {local_file_path}")
        else:
            content = content.replace("[[{{Full_Filename}}|Local File]]", "")
            content = content.replace("![[{{thumbnail}}]]", "")
            logging.warning(f"No local file found for article: {row['Article']} at path: {local_file_path}")

        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)

        logging.info(f"Successfully created note: {filename}")
        return {"success": True, "thumbnail_created": thumbnail_created}
    except Exception as e:
        logging.error(f"Error creating note for {row.get('Article', 'Unknown')}: {str(e)}")
        logging.debug(traceback.format_exc())
        return None


def main():
    input_file = r"G:/Projects/Clwyd Hall/_Resources/Clwyd Hall Project.xlsm"
    output_dir = r"G:/Projects/Obsidian/Vaultez/Newspapers"
    sheet_name = "Newspapers"
    images_dir = os.path.join(output_dir, 'Images')
    articles_dir = os.path.join(output_dir, 'Articles')
    thumbnails_dir = os.path.join(output_dir, 'thumbnails')

    # Create necessary directories
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(images_dir, exist_ok=True)
    os.makedirs(articles_dir, exist_ok=True)
    os.makedirs(thumbnails_dir, exist_ok=True)

    try:
        logging.info(f"Reading Excel file: {input_file}")
        df = pd.read_excel(input_file, sheet_name=sheet_name, engine='openpyxl')
        logging.info(f"Successfully read {len(df)} records from Excel")

        print("Columns in the DataFrame:")
        print(df.columns)

        # Filter rows where Src is NC, BN, or WN
        df_filtered = df[df['Src'].isin(['NC', 'BN', 'WN'])]

        # Update Full_Filename to contain only the filename
        df_filtered['Full_Filename'] = df_filtered['Full_Filename'].apply(os.path.basename)

        notes_created = 0
        thumbnails_created = 0

        for index, row in tqdm(df_filtered.iterrows(), total=len(df_filtered), desc="Processing records"):
            try:
                logging.debug(f"Processing row {index + 1}")
                full_file_path = os.path.join(images_dir, row['Full_Filename'])
                logging.debug(f"Attempting to access file: {full_file_path}")
                result = create_note(row, articles_dir, thumbnails_dir, images_dir)
                if result and result['success']:
                    notes_created += 1
                    if result['thumbnail_created']:
                        thumbnails_created += 1
                else:
                    logging.warning(
                        f"Failed to create note for row {index + 1}. Article: {row.get('Article', 'Unknown')}")
            except Exception as e:
                logging.error(
                    f"Error processing row {index + 1}. Article: {row.get('Article', 'Unknown')}. Error: {str(e)}")
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