# This script generates a series of small images with text overlaid on them,
# then arranges them on an A4-sized PDF for printing.

import json
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import os

# --- Configuration Constants ---
# A4 dimensions in mm
A4_WIDTH_MM = 210
A4_HEIGHT_MM = 297

# Card dimensions in mm (ISO/IEC 7810 ID-1 standard)
CARD_WIDTH_MM = 85.60
CARD_HEIGHT_MM = 53.98

# DPI for image generation. Higher DPI results in better quality but larger files.
DPI = 300
PIXELS_PER_MM = DPI / 25.4

# Convert card dimensions to pixels
CARD_WIDTH_PX = int(CARD_WIDTH_MM * PIXELS_PER_MM)
CARD_HEIGHT_PX = int(CARD_HEIGHT_MM * PIXELS_PER_MM)

# --- File Paths ---
CONFIG_FILE = 'config.json'
DATA_FILE = 'data.xlsx'
OUTPUT_PDF = 'output_cards.pdf'

# --- Functions ---

def load_config(filepath):
    """Loads the text placement configuration from a JSON file."""
    try:
        with open(filepath, 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"Error: Configuration file '{filepath}' not found.")
        return None
    except json.JSONDecodeError:
        print(f"Error: Invalid JSON in configuration file '{filepath}'.")
        return None

def load_data(filepath):
    """Loads data from an Excel file using pandas."""
    try:
        return pd.read_excel(filepath).fillna('')
    except FileNotFoundError:
        print(f"Error: Data file '{filepath}' not found.")
        return None
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def generate_card_image(data_row, config):
    """
    Generates a single card image with text overlaid from the provided data.
    Returns the path to the temporary image file.
    """
    try:
        img = Image.new('RGB', (CARD_WIDTH_PX, CARD_HEIGHT_PX), 'white')
        draw = ImageDraw.Draw(img)

        # Loop through each field defined in the config
        for field, props in config.items():
            text_to_draw = str(data_row.get(field, ''))
            if not text_to_draw:
                continue

            # Position in pixels, converted from mm
            x_pos = int(props['x_mm'] * PIXELS_PER_MM)
            y_pos = int(props['y_mm'] * PIXELS_PER_MM)

            # Define font and color
            try:
                font = ImageFont.truetype(props['font_path'], props['font_size'])
            except IOError:
                # Use a default font if the specified one is not found
                font = ImageFont.load_default()
                print(f"Warning: Font '{props['font_path']}' not found. Using default font.")

            fill_color = tuple(props['color'])

            draw.text((x_pos, y_pos), text_to_draw, font=font, fill=fill_color)

        # Save the generated image to a temporary file
        temp_filename = f"temp_card_{data_row.name}.png"
        img.save(temp_filename)
        return temp_filename
    except Exception as e:
        print(f"Error generating image for data row {data_row.name}: {e}")
        return None

def create_pdf(images):
    """
    Creates a single PDF document and places all generated images on it.
    """
    try:
        c = canvas.Canvas(OUTPUT_PDF, pagesize=A4)

        # Set up a grid for 2 columns and 5 rows
        cols, rows = 2, 5
        # Calculate margins
        x_margin = (A4_WIDTH_MM - (CARD_WIDTH_MM * cols)) / (cols + 1)
        y_margin = (A4_HEIGHT_MM - (CARD_HEIGHT_MM * rows)) / (rows + 1)
        
        # Adjusting the top-left positioning for proper placement
        start_x = x_margin
        start_y = A4_HEIGHT_MM - y_margin - CARD_HEIGHT_MM

        for i, image_path in enumerate(images):
            # Calculate row and column for the current image
            row = i // cols
            col = i % cols
            
            # Calculate coordinates in points for the canvas
            x = (start_x + col * (CARD_WIDTH_MM + x_margin)) * mm
            y = (start_y - row * (CARD_HEIGHT_MM + y_margin)) * mm

            if image_path:
                c.drawImage(image_path, x, y, width=CARD_WIDTH_MM * mm, height=CARD_HEIGHT_MM * mm)

        c.save()
        print(f"Successfully created '{OUTPUT_PDF}' with {len(images)} cards.")
    except Exception as e:
        print(f"Error creating PDF: {e}")
    finally:
        # Clean up temporary image files
        for image_path in images:
            if image_path and os.path.exists(image_path):
                os.remove(image_path)

def main():
    """Main function to orchestrate the entire process."""
    # 1. Load configuration
    config = load_config(CONFIG_FILE)
    if not config:
        return

    # 2. Load data
    data = load_data(DATA_FILE)
    if data is None:
        return

    generated_images = []
    # Limit to a maximum of 10 cards
    num_cards_to_generate = 10

    # 3. Generate individual card images
    for i in range(num_cards_to_generate):
        if i < len(data):
            image_path = generate_card_image(data.iloc[i], config)
            generated_images.append(image_path)
        else:
            # Add a placeholder for a blank card
            # A blank image is still needed to fill the space
            img = Image.new('RGB', (CARD_WIDTH_PX, CARD_HEIGHT_PX), 'white')
            temp_filename = f"temp_blank_card_{i}.png"
            img.save(temp_filename)
            generated_images.append(temp_filename)

    # 4. Create the final PDF
    create_pdf(generated_images)

if __name__ == '__main__':
    main()
