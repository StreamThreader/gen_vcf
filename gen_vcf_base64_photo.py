import base64
import codecs
import sys
from pathlib import Path
from PIL import Image

# version 1.3

if len(sys.argv) == 2:
    photo_file_path = str(sys.argv[1])
else:
    print("use argument as filename, exit...")
    exit()

# reduce photo size
image_file = Image.open(photo_file_path)
image_file = image_file.resize((320,320),Image.Resampling.LANCZOS)
image_file.save(photo_file_path, optimise=True, quality=40)
image_file.close()

# open photo and read data stream
with open(photo_file_path, "rb") as image_file:
    encoded_string = base64.b64encode(image_file.read())
image_file.close()

# file for base64 output
b64_file_path = Path(photo_file_path+".b64")
photo_b64 = codecs.open(b64_file_path, "w", "utf-8")

# read string line character by character
it_firstline = 1
char_count = 0
char_all = 0
for char_obj in encoded_string.decode("utf-8"):
    # check limit excel cell
    char_all += 1
    if char_all == 31000:
        print("limit of excel cell size reached, break")
        print("reduce image size, and try again")
        break

    # if first line, truncate for vcard data
    if char_count == 47 and it_firstline == 1:
        photo_b64.write(char_obj+"\n")
        it_firstline = 0
        char_count = 0
    # if first char in line, add space
    elif char_count == 0 and it_firstline == 0:
        photo_b64.write(" "+char_obj)
        char_count += 1
    # if line reaced 74 symbol, add new line
    elif char_count == 74:
        photo_b64.write(char_obj+"\n")
        char_count = 0
    # if we in middle of line, just add char
    else:
        photo_b64.write(char_obj)
        char_count += 1

photo_b64.close()
print("done")
