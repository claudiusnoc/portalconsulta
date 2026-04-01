from PIL import Image
import os

path = r'assets\icon-guia.png'

if os.path.exists(path):
    img = Image.open(path)
    if img.mode != 'RGBA':
        img = img.convert('RGBA')
    
    bbox = img.getbbox()
    if bbox:
        img_cropped = img.crop(bbox)
        img_cropped.save(path)
        print(f"Successfully cropped! New size: {img_cropped.size}")
    else:
        print("Image is entirely transparent!")
else:
    print(f"File not found: {os.path.abspath(path)}")
