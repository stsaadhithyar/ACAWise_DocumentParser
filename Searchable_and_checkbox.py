import cv2
import numpy as np
from pdf2image import convert_from_path
import os
import img2pdf
import PyPDF2
import shutil

# The function belowis used to rotate the PDF in a specified Degree. 
def roatate_pdf(pdf_path,degree):
    #It takes 2 parameter where the file is opened and the rotated pdf is overwritten in the ame path.
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        pdf_writer = PyPDF2.PdfWriter()

        for page_number in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_number]
            page.rotate(degree)  # Use rotate instead of rotateClockwise
            pdf_writer.add_page(page)

        with open(pdf_path, 'wb') as output_file:
            pdf_writer.write(output_file)


# Function to check if a square is filled based on pixel intensity
def is_square_filled(roi):
    avg_intensity = np.mean(roi)
    return avg_intensity < 200  

# Path to the scanned PDF
pdf_path = "/home/sts852-aadhithyar/Documents/ACA/Main/ACA_Main/rotated.pdf"

#Rotate the PDF Before accessing replacing the characters with balck boxes
roatate_pdf(pdf_path,0) #The parameters are the pdf_path and the degree

# Convert PDF pages to images
images = convert_from_path(pdf_path)

# Directory to save modified images
output_dir = "/home/sts852-aadhithyar/Documents/ACA/Main/ACA_Main/modified_images"
os.makedirs(output_dir, exist_ok=True)

# List to store paths of modified images
modified_image_paths = []

for i, image in enumerate(images):
    # Convert the page image to a NumPy array
    original = np.array(image)

    # Load image, convert to grayscale, Gaussian blur, Otsu's threshold
    gray = cv2.cvtColor(original, cv2.COLOR_BGR2GRAY)
    blur = cv2.GaussianBlur(gray, (3, 3), 0)
    thresh = cv2.threshold(blur, 200, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

    # Find contours
    contours, _ = cv2.findContours(thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)

    # Loop over contours
    for contour in contours:
        # Calculate the area of the contour
        contour_area = cv2.contourArea(contour)
        
        # Filter out contours based on area
        if contour_area < 250:  # Adjust this threshold based on the size of your squares
            continue

        # Approximate the contour to a polygon
        peri = cv2.arcLength(contour, True)
        approx = cv2.approxPolyDP(contour, 0.04 * peri, True)

        # Check if the contour has 4 points (approximates a rectangle)
        if len(approx) == 4:
            # Calculate the bounding rectangle
            x, y, w, h = cv2.boundingRect(contour)
            # Calculate the aspect ratio
            aspect_ratio = w / float(h)
            # Ensure that the aspect ratio is close to 1 (i.e., square)
            if 0.9 <= aspect_ratio <= 1.1:
                # Extract the region of interest (ROI) within the bounding box
                roi = gray[y:y+h, x:x+w]
                # Check if the square is filled
                if is_square_filled(roi):
                    # Draw filled white rectangle around the contour
                    cv2.rectangle(original, (x, y), (x + w, y + h), (0, 0, 0), thickness=cv2.FILLED)

    # Save modified image with bounding boxes
    modified_image_path = os.path.join(output_dir, f"modified_image_{i}.jpg")
    cv2.imwrite(modified_image_path, original)
    modified_image_paths.append(modified_image_path)



# Convert modified images to PDF
pdf_path = "/home/sts852-aadhithyar/Documents/ACA/Main/ACA_Main/GMC1095c1_modified.pdf"
with open(pdf_path, "wb") as f:
    f.write(img2pdf.convert([open(image_path, 'rb') for image_path in modified_image_paths]))

#modified images folder is created which is delted using the below statement.
shutil.rmtree("/home/sts852-aadhithyar/Documents/ACA/Main/ACA_Main/modified_images")
