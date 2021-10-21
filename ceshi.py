import os
from PIL import Image
import glob

img_path = glob.glob("D:/get_image/*.jpg")
path_save = "D:/get_image/"
for file in img_path:
    name = os.path.join(path_save, file)
    im = Image.open(file)
    im.thumbnail((500,500))
    print(im.format, im.size, im.mode)
    im.save(name,'JPEG')



