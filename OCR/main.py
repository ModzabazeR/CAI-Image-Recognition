#!/usr/bin/env python
# -*- coding: utf-8 -*-
from google.cloud import vision
import os
import cv2
from tkinter import filedialog
import moddaeng as md

os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "vision-api-key.json"
client = vision.ImageAnnotatorClient()

image_to_show = filedialog.askopenfilename(filetypes=[("Image File", '.jpg', '.png', '.jpeg'), ("All Files", ".*")])
image_to_scan = md.preprocess(image_to_show)

if image_to_show.endswith('.jpg'):
    content = cv2.imencode('.jpg', image_to_scan)[1].tobytes()
elif image_to_show.endswith('.png'):
    content = cv2.imencode('.png', image_to_scan)[1].tobytes()
elif image_to_show.endswith('.jpeg'):
    content = cv2.imencode('.jpeg', image_to_scan)[1].tobytes()

image = vision.Image(content=content)
response = client.text_detection(image=image)
annotations = response.text_annotations

fontpath = r"C:\Users\modda\AppData\Local\Microsoft\Windows\Fonts\THSarabunNew Bold.ttf"
md.pil_show(image_to_show, annotations, fontpath)