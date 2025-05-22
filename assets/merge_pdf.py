import os
from PyPDF2 import PdfMerger
from reportlab.lib.pagesizes import letter, A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from io import BytesIO
from PIL import Image

def convert_image_to_pdf(image_path, output_pdf_path, page_size=letter, 
                        fit_to_page=True, maintain_aspect_ratio=True, 
                        quality=85, dpi=300):
    """
    Convert a single image to a PDF file with resizing options
    
    Parameters:
    - image_path: Path to the input image
    - output_pdf_path: Path to save the output PDF
    - page_size: PDF page size (default: letter)
    - fit_to_page: Whether to resize image to fit page (default: True)
    - maintain_aspect_ratio: Keep image proportions when resizing (default: True)
    - quality: JPEG quality (1-100) if resampling is needed
    - dpi: Target DPI for the image
    """
    # Open the image with PIL to get details and potentially resize
    img = Image.open(image_path)
    
    # Convert to RGB if needed (ReportLab doesn't support some modes like CMYK directly)
    if img.mode != 'RGB':
        img = img.convert('RGB')
    
    # Calculate target dimensions
    if fit_to_page:
        # Get page dimensions (in points, 1 point = 1/72 inch)
        page_width, page_height = page_size
        
        # Calculate available space (with some margins)
        margin = 20  # 20 points margin on each side
        available_width = page_width - 2 * margin
        available_height = page_height - 2 * margin
        
        # Calculate resize ratio maintaining aspect ratio
        width_ratio = available_width / img.width
        height_ratio = available_height / img.height
        
        if maintain_aspect_ratio:
            resize_ratio = min(width_ratio, height_ratio)
            new_width = int(img.width * resize_ratio)
            new_height = int(img.height * resize_ratio)
        else:
            new_width = int(available_width)
            new_height = int(available_height)
        
        # Resize the image
        img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
        
        # Calculate position to center the image
        x_pos = (page_width - new_width) / 2
        y_pos = (page_height - new_height) / 2
    else:
        # Use original image dimensions
        x_pos, y_pos = 0, 0
        new_width, new_height = img.size
    
    # Create a temporary buffer for the PDF
    buffer = BytesIO()
    
    # Create PDF with the specified page size
    c = canvas.Canvas(buffer, pagesize=page_size)
    
    # Save the resized image to a temporary buffer
    temp_img_buffer = BytesIO()
    img.save(temp_img_buffer, format='JPEG', quality=quality, dpi=(dpi, dpi))
    temp_img_buffer.seek(0)
    
    # Draw the image on the PDF
    c.drawImage(ImageReader(temp_img_buffer), x_pos, y_pos, 
                width=new_width, height=new_height, 
                preserveAspectRatio=maintain_aspect_ratio)
    c.showPage()
    c.save()
    
    # Write the buffer to a file
    with open(output_pdf_path, 'wb') as f:
        f.write(buffer.getvalue())

def merge_all_to_pdf(input_folder, output_pdf, 
                    page_size=letter, fit_to_page=True, 
                    maintain_aspect_ratio=True, quality=85, dpi=300):
    """Merge all PDFs and images in a folder into a single PDF with resizing options"""
    merger = PdfMerger()
    temp_files = []
    
    # Supported image extensions
    image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff']
    
    try:
        # Process all files in the input folder
        for filename in sorted(os.listdir(input_folder)):
            filepath = os.path.join(input_folder, filename)
            
            if filename.lower().endswith('.pdf'):
                # Add PDF files directly
                merger.append(filepath)
            elif any(filename.lower().endswith(ext) for ext in image_extensions):
                # Convert images to temporary PDFs with resizing
                temp_pdf = f"temp_{filename}.pdf"
                convert_image_to_pdf(
                    filepath, temp_pdf,
                    page_size=page_size,
                    fit_to_page=fit_to_page,
                    maintain_aspect_ratio=maintain_aspect_ratio,
                    quality=quality,
                    dpi=dpi
                )
                merger.append(temp_pdf)
                temp_files.append(temp_pdf)
        
        # Write the merged PDF to the output file
        merger.write(output_pdf)
        print(f"Successfully created merged PDF: {output_pdf}")
        
    except Exception as e:
        print(f"Error occurred: {str(e)}")
    finally:
        # Clean up temporary files
        merger.close()
        for temp_file in temp_files:
            try:
                os.remove(temp_file)
            except:
                pass

if __name__ == "__main__":
    # Configuration - change these as needed
    input_directory = "Experience"  # Folder containing PDFs and images
    output_pdf_file = input_directory + ".pdf"  # Output PDF filename
    
    # Image conversion settings
    settings = {
        'page_size': A4,  # Options: letter, A4, etc. or (width, height) in points
        'fit_to_page': True,  # Resize images to fit the PDF page
        'maintain_aspect_ratio': True,  # Keep original proportions when resizing
        'quality': 90,  # JPEG quality (1-100)
        'dpi': 300  # Target DPI for images
    }
    
    # Create input directory if it doesn't exist
    if not os.path.exists(input_directory):
        os.makedirs(input_directory)
        print(f"Created input directory: {input_directory}")
        print("Please add your PDF and image files to this directory and run the script again.")
    else:
        merge_all_to_pdf(input_directory, output_pdf_file, **settings)