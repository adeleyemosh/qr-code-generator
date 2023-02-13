import os
from PIL import Image
import qrcode
from PIL import Image, ImageDraw, ImageFont

def generate_qr_with_label(prefix, start, end, step=1):
    # Determine the directory name based on the prefix
    if 'AEDC' in prefix:
        directory = 'QRCodes/AEDC'
    elif 'ECG' in prefix:
        directory = 'QRCodes/ECG'
    else:
        directory = 'QRCodes'

    # Create the directory if it doesn't exist
    os.makedirs(directory, exist_ok=True)

    # Create an empty image to hold the QR Code and label
    qr_size = 140  # Set the size of the QR Code image
    font_size = 14
    font_path = "C:\\Users\Moshood\\OneDrive\\dev\\python_data_profiling_scripts\\qrcode_generation\\fonts\\MonoBold-z8jG0.ttf"
    font = ImageFont.truetype(font_path, font_size, encoding="unic")

    for num in range(start, end+1, step):
        # Create a new image to hold the QR code and label for the current tag
        code = prefix + str(num).zfill(5)  # Pad the number with zeros to 5 digits
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
