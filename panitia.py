import os

import img2pdf
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

xls = pd.ExcelFile('src/uxo.xlsx')

title_font = ImageFont.truetype('src/font/Niconne-Regular.ttf', 120)

for sheets in xls.sheet_names:
    img_path = f"data/panitia/{sheets}"
    pdf_path = f"sertifikat/panitia/{sheets}"
    if not os.path.exists(img_path):
        os.makedirs(img_path)
    if not os.path.exists(pdf_path):
        os.makedirs(pdf_path)

    df = pd.read_excel(xls, sheets, skiprows=3, usecols="B")
    for (i, nama) in df.iterrows():
        image = Image.open("src/panitia.png")
        W, H = image.size

        image_editable = ImageDraw.Draw(image)
        text = nama[0].title()
        nfile = text.replace(" ", "-")

        _, _, w, h = image_editable.textbbox((0, 0), text, font=title_font)
        image_editable.text(((W-w)/2, ((H-h)/2)-40), text,
                            (39, 79, 115), font=title_font)
        image.save(f"{img_path}/{nfile}.png")
        with open(f"{pdf_path}/{nfile}.pdf", "wb") as f:
            f.write(img2pdf.convert(f"{img_path}/{nfile}.png"))
            print(f"{pdf_path}/{nfile}.pdf")
