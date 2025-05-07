import os
import pandas as pd
import logging
import sys
import time
from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfReader, PdfWriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Setup logging
def setup_logging(log_path):
    """Set up logging to file and console."""
    log_format = '%(asctime)s - %(levelname)s - %(message)s'
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        handlers=[
            logging.FileHandler(log_path),
            logging.StreamHandler(sys.stdout)
        ]
    )
    return logging.getLogger()

# Create output directory
def create_output_directory(output_dir):
    """Create output directory if it doesn't exist."""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    return output_dir

# Load Excel data
def load_excel_data(excel_path, test_mode=False):
    """Load data from Excel file."""
    logger.info(f"Loading data from {excel_path}")
    try:
        df = pd.read_excel(excel_path)
        
        # Print available columns to help diagnose
        logger.info(f"Available columns: {', '.join(df.columns)}")
        
        # Use more flexible column name matching (case insensitive)
        id_col = next((col for col in df.columns if col.lower() in ["id", "participant id", "student id"]), None)
        name_col = next((col for col in df.columns if "name" in col.lower()), None)
        email_col = next((col for col in df.columns if "email" in col.lower()), None)
        
        if not id_col or not name_col:
            logger.error(f"Could not find required columns. Need ID and name columns.")
            logger.error(f"Found: ID column: {id_col}, Name column: {name_col}")
            sys.exit(1)
            
        # Rename columns to match expected names
        column_mapping = {}
        if id_col: column_mapping[id_col] = "ID"
        if name_col: column_mapping[name_col] = "Full Name"
        if email_col: column_mapping[email_col] = "Student Email"
        
        df = df.rename(columns=column_mapping)
        
        if test_mode:
            logger.info("Test mode: Using first 5 entries only")
            return df.head(5)
        return df
    except Exception as e:
        logger.error(f"Failed to load Excel file: {str(e)}")
        sys.exit(1)

def create_certificate_pdf(template_pdf_path, output_dir, row):
    """Create certificate by adding text directly to PDF."""
    try:
        # Create output filename
        pdf_path = os.path.join(output_dir, f"{row['ID']}.pdf")
        
        # Register the Bradley Hand font if available
        try:
            # Try to find Bradley Hand font (locations vary by system)
            possible_font_paths = [
                '/Library/Fonts/Bradley Hand Bold.ttf',
                '/Library/Fonts/Bradley Hand.ttf',
                '/Library/Fonts/BradleyHandITCTT-Bold.ttf'
            ]
            
            font_found = False
            for font_path in possible_font_paths:
                if os.path.exists(font_path):
                    pdfmetrics.registerFont(TTFont('BradleyHand', font_path))
                    font_name = 'BradleyHand'
                    font_found = True
                    break
                    
            if not font_found:
                # Fallback to a standard font
                font_name = 'Helvetica-Bold'
                logger.warning("Bradley Hand font not available, using Helvetica Bold instead")
        except Exception as e:
            # Fallback to a standard font
            font_name = 'Helvetica-Bold'
            logger.warning(f"Font registration error: {str(e)}. Using Helvetica Bold instead")
        
        # Step 1: Create a PDF with just the name
        temp_overlay = os.path.join(output_dir, f"overlay_{row['ID']}.pdf")
        c = canvas.Canvas(temp_overlay, pagesize=landscape(A4))
        c.setFont(font_name, 24)
        
        # Get text width to calculate center position
        full_name = row['Full Name']
        text_width = c.stringWidth(full_name, font_name, 20)
        
        # Get page dimensions (in points)
        page_width, page_height = landscape(A4)
        
        # Calculate center position
        x_center = page_width * 0.5 - (text_width / 2)
        
        # Lower the vertical position - adjust this value as needed
        # A4 landscape is 841.89 x 595.28 points
        # Try different values like 250-350 for y_position
        y_position = 280  # Lower value moves text down
        
        # Debug logging to help with positioning
        logger.info(f"Page width: {page_width}, Text width: {text_width}")
        logger.info(f"Placing text at x: {x_center}, y: {y_position}")
        
        # Draw the name centered horizontally
        c.drawString(x_center, y_position, full_name)
        c.save()
        
        # Step 2: Merge the name overlay with the template
        template_pdf = PdfReader(open(template_pdf_path, 'rb'))
        overlay_pdf = PdfReader(open(temp_overlay, 'rb'))
        
        output = PdfWriter()
        page = template_pdf.pages[0]
        page.merge_page(overlay_pdf.pages[0])
        output.add_page(page)
        
        with open(pdf_path, "wb") as output_stream:
            output.write(output_stream)
        
        # Clean up temporary file
        if os.path.exists(temp_overlay):
            os.remove(temp_overlay)
            
        logger.info(f"Created certificate PDF for ID {row['ID']}")
        return True, None
        
    except Exception as e:
        error_msg = f"PDF creation error: {str(e)}"
        logger.error(error_msg)
        return False, error_msg

def process_certificates(excel_path, template_pdf_path, output_dir, test_mode=False):
    """Process all certificates."""
    start_time = time.time()
    
    # Load data
    df = load_excel_data(excel_path, test_mode)
    
    # Process records
    total = len(df)
    successful = 0
    failed = 0
    errors = []
    
    logger.info(f"Starting to process {total} certificates")
    
    for index, row in df.iterrows():
        logger.info(f"Processing {index+1}/{total}: ID {row['ID']}, Name: {row['Full Name']}")
        success, error = create_certificate_pdf(template_pdf_path, output_dir, row)
        
        if success:
            successful += 1
        else:
            failed += 1
            errors.append(f"ID {row['ID']}: {error}")
            logger.error(f"Failed to process ID {row['ID']}: {error}")
    
    # Log summary
    elapsed_time = time.time() - start_time
    logger.info(f"Completed processing {total} certificates in {elapsed_time:.2f} seconds")
    logger.info(f"Successful: {successful}, Failed: {failed}")
    
    if failed > 0:
        logger.info("Error summary:")
        for error in errors:
            logger.info(f"  {error}")
    
    return successful, failed, errors

def main():
    # Configuration
    excel_path = "IC-Juniors-Participants.xlsx"
    
    template_pdf_path = "Imagine Cup Junior Participation-2025.pdf"
    # output_dir = "pdf_certificates_output"
    
    output_dir = 'certificates'
    
    # Create output directory
    create_output_directory(output_dir)
    
    # Setup logging
    log_path = os.path.join(output_dir, "pdf_certificate_generation.log")
    global logger
    logger = setup_logging(log_path)
    
    logger.info("PDF Certificate Generation Process Started")
    logger.info("====================================")
    
    # Ask if this is a test run
    test_mode = input("Run in test mode (first 5 entries only)? (y/n): ").lower() == 'y'
    
    # Process certificates
    successful, failed, errors = process_certificates(
        excel_path, template_pdf_path, output_dir, test_mode=test_mode
    )
    
    if test_mode and successful > 0:
        proceed = input("Test completed. Proceed with full processing? (y/n): ").lower() == 'y'
        if proceed:
            logger.info("Starting full processing...")
            process_certificates(excel_path, template_pdf_path, output_dir, test_mode=False)
    
    logger.info("PDF Certificate Generation Process Completed")
    logger.info("====================================")

if __name__ == "__main__":
    main()