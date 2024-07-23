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

# Set up extensive logging
logging.basicConfig(filename='newspaper_import_log.txt', level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Define the template as a string
TEMPLATE = """---
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

# {{Article}}

## Source Information
- Newspaper: {{Newspaper_or_Source}}
- Published: {{Published}}

## People Involved
{{people_involved}}

## Locations
{{locations}}

## Links
- [Online Source]({{Web}})
- [[{{Full_Filename}}|Local File]]

## Thumbnail
![[{{thumbnail}}]]

## Tags
#Llanychan #Taber-Project #{{Src}}
{{tags}}
"""


def clean_filename(filename):
    # Remove any characters that are not alphanumeric, space, comma, or period
    cleaned = re.sub(r'[^\w\s,.-]', '', filename)
    # Replace multiple spaces with a single space
    cleaned = re.sub(r'\s+', ' ', cleaned)
    return cleaned.strip()


def create_thumbnail(file_path, output_dir, size=(300, 300)):
    try:
        # Get the base name without extension
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        thumbnail_filename = f"thumb_{base_name}.jpg"
        thumbnail_path = os.path.join(output_dir, 'thumbnails', thumbnail_filename)

        if file_path.lower().endswith('.pdf'):
            # Handle PDF files
            doc = fitz.open(file_path)
            pix = doc[0].get_pixmap()
            pix.save(thumbnail_path)
            # Resize the saved image
            with Image.open(thumbnail_path) as img:
                img.thumbnail(size)
                img.save(thumbnail_path)
        else:
            # Handle image files
            with Image.open(file_path) as img:
                # Convert palette mode (P) to RGB
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


def create_note(row, output_dir):
    try:
        logging.info(f"Starting to create note for article: {row['Article']}")

        # Create filename
        filename = f"{clean_filename(row['Article'])}.md"
        if len(filename) > 255:
            filename = filename[:252] + '.md'
        filepath = os.path.join(output_dir, filename)

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
        if pd.notna(row['Name_1']):
            people_involved.append(f"- {row['Name_1']}")
        if pd.notna(row['Name_2']):
            people_involved.append(f"- {row['Name_2']}")
        content = content.replace("{{people_involved}}", "\n".join(people_involved))

        # Process locations
        locations = []
        if pd.notna(row['Place_1']):
            locations.append(f"- {row['Place_1']}")
        if pd.notna(row['Place_2']):
            locations.append(f"- {row['Place_2']}")
        content = content.replace("{{locations}}", "\n".join(locations))

        # Process tags
        tags = []
        if pd.notna(row['T']):
            tags.append(f"#Theme-{row['T']}")
        for i in range(2, 6):
            if pd.notna(row[f'Theme_{i}']):
                tags.append(f"#Theme-{row[f'Theme_{i}']}")
        for i in range(1, 3):
            if pd.notna(row[f'Name_{i}']):
                tags.append(f"#Person-{row[f'Name_{i}']}")
        for i in range(1, 3):
            if pd.notna(row[f'Place_{i}']):
                tags.append(f"#Place-{row[f'Place_{i}']}")
        content = content.replace("{{tags}}", " ".join(tags))

        # Handle thumbnail
        thumbnail_created = False
        local_file_path = row['Full_Filename']
        if pd.notna(local_file_path) and os.path.exists(local_file_path):
            thumbnail_result = create_thumbnail(local_file_path, output_dir)
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

        # Update Full_Filename
        df_filtered['Full_Filename'] = 'G:\\Projects\\_Resources\\Newspapers\\images\\' + df_filtered['Full_Filename']

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
