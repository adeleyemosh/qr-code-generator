import os
from PIL import Image
import csv
import qrcode
import openpyxl
from PIL import Image, ImageDraw, ImageFont

def generate_qr_with_label(prefix, start, end, step=1):
    # Determine the directory name based on the prefix
    if 'AEDC' in prefix:
        directory = 'QRCodes/AEDC'
    elif 'ECG' in prefix:
        directory = 'QRCodes/ECG'
    elif 'YEDC' in prefix:
        directory = 'QRCodes/YEDC'
    elif 'IEDC' in prefix:
        directory = 'QRCodes/IEDC'
    else:
        directory = 'QRCodes'

    # Create the directory if it doesn't exist
    os.makedirs(directory, exist_ok=True)

    # Create an empty image to hold the QR Code and label
    qr_size = 140  # Set the size of the QR Code image
    font_size = 14
    font_path = os.path.join(os.path.expanduser("~"),"OneDrive", "dev", "qr_code_tag_generator", "fonts", "MonoBold-z8jG0.ttf")
    font = ImageFont.truetype(font_path, font_size, encoding="unic")

    # start = int(start)
    # end = int(end)
    # step = int(step)
    for num in range(start, end+1, step):
        # Create a new image to hold the QR code and label for the current tag
        code = prefix + str(num).zfill(7)  # Pad the number with zeros to 5 digits
        img = Image.new('RGB', (qr_size, qr_size+font_size), color='white')

        # Generate the QR Code for the current tag and add it to the image
        qr = qrcode.QRCode(box_size=5)
        qr.add_data(code)
        qr_img = qr.make_image()
        qr_width, qr_height = qr_img.size
        qr_x = (qr_size - qr_width) // 2  # Center the QR Code image horizontally
        img.paste(qr_img, (qr_x, 0))

        # Add the label text below the QR Code
        label_text = code
        label_width, label_height = font.getsize(label_text)
        label_x = (qr_size - label_width) // 2  # Center the label text horizontally
        label_y = qr_height - 15  # Place the label below the QR Code
        label_draw = ImageDraw.Draw(img)
        label_draw.text((label_x, label_y), label_text, font=font, fill='black', align='center',
                        spacing=1, stroke_width=1, stroke_fill='white', antialias=True
                        )

        # Save the QR code and label as a PNG file with the code as the file name
        filename = os.path.join(directory, code + '.png')
        img.save(filename)

    return

def generate_qr_code(code, file_name):
    qr_size = 140  # Set the size of the QR Code image
    font_size = 14
    font_path = os.path.join(os.path.expanduser("~"),"OneDrive", "dev", "qr_code_tag_generator", "fonts", "MonoBold-z8jG0.ttf")
    font = ImageFont.truetype(font_path, font_size, encoding="unic")
    img_size = 150

    # Generate the QR Code for the current tag and add it to the image
    qr = qrcode.QRCode(box_size=5)
    qr.add_data(code)
    qr_img = qr.make_image()
    qr_width, qr_height = qr_img.size
    qr_x = (qr_size - qr_width) // 2  # Center the QR Code image horizontally
    img = Image.new('RGB', (img_size, img_size), color='white')
    img.paste(qr_img, (qr_x, 0))

    # Add the label text below the QR Code
    label_text = file_name
    label_width, label_height = font.getsize(label_text)
    label_x = (qr_size - label_width) // 2  # Center the label text horizontally
    label_y = qr_height - 15  # Place the label below the QR Code
    label_draw = ImageDraw.Draw(img)
    label_draw.text((label_x, label_y), label_text, font=font, fill='black', align='center',
                    spacing=1, stroke_width=1, stroke_fill='white', antialias=True
                    )

    return img

def generate_qr_from_file(prefix, file_path):
    if file_path.endswith('.csv'):
        with open(file_path, 'r') as csv_file:
            reader = csv.reader(csv_file)
            numbers = [f"{prefix}{n}" for row in reader for n in row if n is not None]
    elif file_path.endswith(('.xlsx', '.xlsm', '.xltx', '.xltm')):
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        numbers = [f"{prefix}{cell.value}" for row in ws.iter_rows() for cell in row if cell.value is not None]
    else:
        raise ValueError("Invalid file type. Must be CSV or Excel file.")

    # Determine the directory name based on the prefix
    if 'AEDC' in prefix:
        img_directory = 'QRCodes/AEDC'
    elif 'ECG' in prefix:
        img_directory = 'QRCodes/ECG'
    elif 'YEDC' in prefix:
        img_directory = 'QRCodes/YEDC'
    elif 'IEDC' in prefix:
        img_directory = 'QRCodes/IEDC'
    else:
        img_directory = 'QRCodes'

    # Create the directory if it doesn't exist
    os.makedirs(img_directory, exist_ok=True)

    qr_codes = []
    for num in numbers:
        img = generate_qr_code(num, num)
        img_name = f'{num}.png'
        img_path = os.path.join(img_directory, img_name)
        img.save(img_path)
        qr_codes.append(img_path)

    print(f"QR codes generated for {len(qr_codes)} numbers")

    return qr_codes