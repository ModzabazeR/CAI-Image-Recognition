from PIL import Image, ImageOps, ImageFont, ImageDraw
import os
import cv2
import imutils
import numpy as np

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

def save_text(annots):
    annots = annots[:1]
    with open('output.txt', 'w', encoding='utf-8') as f:
        for text in annots:
            f.write('{}\n'.format(text.description))