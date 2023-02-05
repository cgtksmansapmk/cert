import os

import gspread
import img2pdf
from PIL import Image, ImageDraw, ImageFont

gc = gspread.oauth(credentials_filename='src/credentials.json',
                   authorized_user_filename='src/authorized_user.json')
sprit = gc.open_by_url(
    'https://docs.google.com/spreadsheets/d/1t3YVKUvAGQxmGyJsNo0MCfMQcMWc98bsh8hGiDI6sLE')

title_font = ImageFont.truetype('src/font/Niconne-Regular.ttf', 120)


# loop semua sheet dan ambil sheet offering saja
for sheet in sprit.worksheets()[5:]:
    sheet = str(sheet).split(" ")[1].replace("'", "")
    shit = sprit.worksheet(sheet)

    # buat folder untuk sheet offering
    img_path = f"data/peserta/{sheet}"
    pdf_path = f"sertifikat/peserta/{sheet}"
    if not os.path.exists(img_path):
        os.makedirs(img_path)
    if not os.path.exists(pdf_path):
        os.makedirs(pdf_path)

    # loop semua nama dalam sheet offering
    for nama in shit.col_values(1):
        image = Image.open("src/peserta.png")
        W, H = image.size

        image_editable = ImageDraw.Draw(image)
        text = nama.title()
        nfile = text.replace(" ", "-")

        _, _, w, h = image_editable.textbbox((0, 0), text, font=title_font)
        image_editable.text(((W-w)/2, ((H-h)/2)-40), text,
                            (39, 79, 115), font=title_font)
        image.save(f"{img_path}/{nfile}.png")
        with open(f"{pdf_path}/{nfile}.pdf", "wb") as f:
            f.write(img2pdf.convert(f"{img_path}/{nfile}.png"))
            print(f"{pdf_path}/{nfile}.pdf")
