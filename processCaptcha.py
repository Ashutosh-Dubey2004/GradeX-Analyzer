import cv2
import numpy as np
import easyocr
from PIL import Image
from io import BytesIO

reader = easyocr.Reader(["en"])

def process_captcha(image_element):
    """Processes captcha using EasyOCR."""
    ''' 
    image = Image.open(BytesIO(image_element.screenshot_as_png))
    image_np = np.array(image)

    # Convert to grayscale
    gray = cv2.cvtColor(image_np, cv2.COLOR_RGB2GRAY)

    # Use EasyOCR to extract text
    extracted_text = reader.readtext(gray, detail=0)
    
    # Join detected text and remove spaces
    captcha_text = "".join(extracted_text).replace(" ", "")
    return captcha_text
    '''
    
    return "".join(reader.readtext(cv2.cvtColor(np.array(Image.open(BytesIO(image_element.screenshot_as_png))), cv2.COLOR_RGB2GRAY), detail=0)).replace(" ", "")
