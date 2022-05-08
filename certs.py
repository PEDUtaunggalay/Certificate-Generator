#Imports Required Packages from PIL
from PIL import Image, ImageDraw, ImageFont

#Import Pandas for better access of Data and .xlsx File
import pandas as pd

#Import the file that contains all the details
data = pd.read_excel("fb.xlsx")

#Import 'Name' List from the imported .xlsx file
name_list = data['Name'].to_list()

#Import 'Father' List from the imported .xlsx file
father_list = data['Father'].to_list()

id_list = data['Phlon No.'].to_list()

#Determining the length of the list
max_no = len(name_list)

#The Loops for creating certificates in bulk
# for i in name_list:
for idx, i in enumerate(name_list):
    # background = Image.open("certa4.jpg")
    # overlay = Image.open("1.png")
    # background = background.convert("RGBA")
    # overlay = overlay.convert("RGBA")
    # new_img = Image.blend(background, overlay, 0.5)

    # getting student id 
    id = id_list[idx]
    print(idx)
    print(id)
    
    # opening certificate template
    img1 = Image.open(r"certa4.jpg")

    # opening student's photo 
    img2 = Image.open("%d.png" % (id))

    # No transparency mask specified, 
    # simulating an raster overlay
    img1.paste(img2, (250,500))
    # im = img1.convert("RGBA")
    # img1 = Image.blend(img1, img2, 0.5)
    # img1.show()
    # img1.save("test.png")
    # img1 = img1.convert("RGBA")

    # no need to reopen another image, so?
    # im = Image.open(img1)
    father = father_list[idx]
    # im = Image.open("certa4.jpg")
    d = ImageDraw.Draw(img1)
    location = (275, 1050)
    location2 = (275,1250)
    text_color = (0, 137, 209)
    font = ImageFont.truetype("fonts/AlexBrush-Regular.ttf", 250, encoding="unic")
    d.text(location, i.title(), fill=text_color, font=font)
    d.text(location2, father.title(), fill=text_color, font=font)
    
    img1.save("certificate_"+i+".pdf")
    print("(%d/%d) Certificate Created for:  %s" % (idx+1, max_no, i.title()))
    
    
print("""\n*************************
All Certificates Created! \nPlease follow me on Github for more awesome codes\n\tGitHub: https://github.com/mursalfk
*************************
""")
#Read readme.md for further instructions