from PIL import Image, ImageOps, ImageFont, ImageDraw
from google.cloud.vision import AnnotateImageResponse
import os
import cv2
import imutils
import numpy as np
import json
import pdfplumber

# ------------ Global Tools ------------ #

def pretty_save_json(file: str, data: dict) -> None:
    with open(file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def to_float(value: str):
    "Convert string to float"
    return float(value.replace(',', ''))

def to_string(value: float):
    "Convert float to string with 2 decimal places"
    return '{:,.2f}'.format(value)

# ------------ PDF Tools ------------ #

def ie_extract_text(path: str) -> str:
    text = ""
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text += page.extract_text()
    return text

# ------------ Image Preprocessing and debuging for Google Vision API ------------ #

def picture_correction(image_path: str):
    "Rotate image if necessary and save it to 300 dpi"
    pil_img = Image.open(image_path)
    pil_img = ImageOps.exif_transpose(pil_img)
    pil_img.save(f"temp300_{os.path.basename(image_path)}", dpi=(300, 300))

def preprocess(image_path: str) -> np.ndarray:
    "Preprocess image for OCR"
    rect_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (70, 15))
    picture_correction(image_path)

    # Image Binarization
    image = cv2.imread(f"temp300_{os.path.basename(image_path)}")
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blackhat = cv2.morphologyEx(gray, cv2.MORPH_BLACKHAT, rect_kernel)

    os.remove(f"temp300_{os.path.basename(image_path)}")
    return blackhat

def cv_show(img: str, annots):
    "Show detected text on image with OpenCV (Thai not supported)"
    #skip 1st item
    annots = annots[1:]
    img = cv2.imread(img)
    print('Texts:')
    for text in annots:
        # print('\n"{}"'.format(text.description))

        vertices = (['({},{})'.format(vertex.x, vertex.y)
                    for vertex in text.bounding_poly.vertices])

        # print('bounds: {}'.format(','.join(vertices)))
        cv2.rectangle(img, (text.bounding_poly.vertices[0].x, text.bounding_poly.vertices[0].y),
                    (text.bounding_poly.vertices[2].x, text.bounding_poly.vertices[2].y), (0, 255, 0), 2)
        cv2.putText(img, text.description, (text.bounding_poly.vertices[0].x, text.bounding_poly.vertices[0].y),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 0, 255), 2)

    cv2.imwrite('result_cv.jpg', img)
    cv2.imshow('img', imutils.resize(img, height=800))
    cv2.waitKey(0)

def pil_show(img: str, annots, fontpath: str):
    "Show detected text on image with Pillow (Thai supported)"
    font = ImageFont.truetype(fontpath, 24)
    annots = annots[1:]

    img_pil = Image.open(img)
    img_pil = ImageOps.exif_transpose(img_pil)
    draw = ImageDraw.Draw(img_pil)

    for text in annots:
        # print('\n"{}"'.format(text.description))
        vertices = (['({},{})'.format(vertex.x, vertex.y)
                    for vertex in text.bounding_poly.vertices])
        # print('bounds: {}'.format(','.join(vertices)))
        draw.rectangle([(text.bounding_poly.vertices[0].x, text.bounding_poly.vertices[0].y),
                        (text.bounding_poly.vertices[2].x, text.bounding_poly.vertices[2].y)], outline=(0, 255, 0), width=4)
        draw.text((text.bounding_poly.vertices[0].x, text.bounding_poly.vertices[0].y - 25), text.description, font=font, fill=(0, 0, 255))

    img = np.array(img_pil)
    cv2.imwrite('result_pil.jpg', img)
    cv2.imshow('img', imutils.resize(img, height=800))
    cv2.waitKey(0)

def save_pil_img(output_path:str, img: str, annots, fontpath: str):
    "Save image with detected text to output path"
    font = ImageFont.truetype(fontpath, 24)
    annots = annots[1:]

    img_pil = Image.open(img)
    img_pil = ImageOps.exif_transpose(img_pil)
    draw = ImageDraw.Draw(img_pil)

    for text in annots:
        draw.rectangle([(text.bounding_poly.vertices[0].x, text.bounding_poly.vertices[0].y),
                        (text.bounding_poly.vertices[2].x, text.bounding_poly.vertices[2].y)], outline=(0, 255, 0), width=4)
        draw.text((text.bounding_poly.vertices[0].x, text.bounding_poly.vertices[0].y - 25), text.description, font=font, fill=(0, 0, 255))

    img = np.array(img_pil)
    cv2.imwrite(output_path, img)

# ------------  Google Vision API Tools ------------ #

def save_text(annots):
    annots = annots[:1]
    with open('output.txt', 'w', encoding='utf-8') as f:
        for text in annots:
            f.write('{}\n'.format(text.description))

def get_content(image_path, cv_image):
    "Get content from numpy array image and convert to bytes to be processed by Google Vision API"
    supported_format = ('.jpg', '.png', '.jpeg')
    for format in supported_format:
        if image_path.endswith(format):
            return cv2.imencode(format, cv_image)[1].tobytes()
    else:
        raise ValueError('Unsupported image format')

def get_json_response(raw_response):
    "Get json string from response"
    serialized_proto_plus = AnnotateImageResponse.serialize(raw_response)
    response_json = AnnotateImageResponse.deserialize(serialized_proto_plus)
    response_json = AnnotateImageResponse.to_json(response_json)
    return response_json
