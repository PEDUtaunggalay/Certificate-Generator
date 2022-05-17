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

#Import 'Phlon No.' List from the imported .xlsx file
id_list = data['Phlon No.'].to_list()

#Import 'Gender' List from the imported .xlsx file
sex_list = data['Sex'].to_list()

#Determining the length of the list
max_no = len(name_list)

#The Loops for creating certificates in bulk
# for i in name_list:
for idx, i in enumerate(name_list):

    # https://stackoverflow.com/a/10640168/5307753 
    # background = Image.open("certa4.jpg")
    # overlay = Image.open("1.png")
    # background = background.convert("RGBA")
    # overlay = overlay.convert("RGBA")
    # new_img = Image.blend(background, overlay, 0.5)

    # getting student id from id_list which is imported from .xlsx file
    id = id_list[idx]
    print(id)
    fm = sex_list[idx]
    print(fm)
    
    # https://www.geeksforgeeks.org/overlay-an-image-on-another-image-in-python/
    # opening certificate template
    img1 = Image.open(r"pbm-2022-template.jpg") # certa4.jpg
    # img1.show()
    # opening student's photo 
    img2 = Image.open("images/photos/pbm-2022/%d.png" % (id))
    # No transparency mask specified, 
    # simulating an raster overlay
    img1.paste(img2, (2760,815))
    # im = img1.convert("RGBA")
    # img1.show()
    # img1.save("test.png")

    # getting father's name
    father = father_list[idx]

    # getting student's gender
    sex = sex_list[idx]

    # im = Image.open("certa4.jpg")
    # getting ready to start writing names on certificate
    d = ImageDraw.Draw(img1)
    # location for student name 
    location = (350, 925)

    # location2 for gender identity 
    location2 = (293, 1175)

    # location for father's name 
    location3 = (300,1300)

    # location for Phlon No. 
    location4 = (2875,1420)

    # blue 
    # text_color = (0, 137, 209)
    text_color = (0, 0, 0)
    text_color2 = (9,35,138)

    # orignal font fonts/AlexBrush-Regular.ttf 
    font = ImageFont.truetype("fonts/SnellRoundhand/SnellBT-Bold.otf", 150, encoding="unic")
    # font2 = ImageFont.truetype("fonts/AlexBrush-Regular.ttf", 250, encoding="unic")
    font2 = ImageFont.truetype("fonts/AvenirNext/AvenirNextLTPro-Regular.otf", 70, encoding="unic")

    font3 = ImageFont.truetype("fonts/AvenirNext/AvenirNextLTPro-Bold.otf", 60, encoding="unic")

    # draw student's name
    d.text(location, i.title(), fill=text_color, font=font)

    # draw female/male 
    if sex.title() == "Male":
        d.text(location2, "Son of", fill=text_color2, font=font2)
    else: 
        d.text(location2, "Daughter of", fill=text_color2, font=font2)
    
    # father's name 
    d.text(location3, father.title(), fill=text_color, font=font)

    # draw Phlon No. 
    d.text(location4, "P-No.%d" % id, fill=text_color, font=font3)

    # img1.show()
    img1.save("pbm_2022_b1_%d.jpg" % (id))
    print("(%d/%d) Certificate Created for:  %s" % (idx+1, max_no, i.title()))
    
    
print("""\n*************************
All Certificates Created! \nPlease follow me on Github for more awesome codes\n\tGitHub: https://github.com/mursalfk
*************************
""")
#Read readme.md for further instructions