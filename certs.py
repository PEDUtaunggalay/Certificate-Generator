#Imports Required Packages from PIL
from PIL import Image, ImageDraw, ImageFont

#Import Pandas for better access of Data and .xlsx File
import pandas as pd

#Import the file that contains all the details
data = pd.read_excel("pbm_2022_b1.xlsx")

#Import 'Name' List from the imported .xlsx file
name_list = data['Name'].to_list()

#Import 'Father' List from the imported .xlsx file
father_list = data['Father'].to_list()

#Import 'Phlon No.' List from the imported .xlsx file
phlon_id_list = data['Phlon No.'].to_list()

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
    phlon_id = phlon_id_list[idx]

    # “how to get a zero before 1-9 in python”
    # https://www.codegrepper.com/code-examples/python/how+to+get+a+zero+before+1-9+in+python
    str(phlon_id).zfill(4)
    print(str(phlon_id).zfill(4))
    id = str(phlon_id).zfill(4)
    
    # https://www.geeksforgeeks.org/overlay-an-image-on-another-image-in-python/
    
    # opening certificate template
    img1 = Image.open(r"images/pbm-2022-template.jpg") # certa4.jpg
    # img1.show()
    
    # opening student's photo 
    img2 = Image.open("images/dhammapada-2024/photos/%s.png" % (id))

    new_image = img2.resize((554, 554))

    # adding student's photo
    # No transparency mask specified, 
    # simulating an raster overlay
    img1.paste(new_image, (2760,815))
    # im = img1.convert("RGBA")
    # img1.show()
    # img1.save("test.png")

    # getting father's name
    father = father_list[idx]

    # getting student's gender
    sex = sex_list[idx]
    print(sex)

    # im = Image.open("certa4.jpg")
    # getting ready to start writing names on certificate
    d = ImageDraw.Draw(img1)

    # location for student name 
    location_student_name = (350, 925)

    # location2 for daughter of/son of
    location_daughter_son = (293, 1175)

    # location for father's name 
    location_father_name = (300,1300)

    # location for Phlon No. 
    location_p_no = (2875,1420)

    # blue 
    # text_color = (0, 137, 209)
    text_color = (0, 0, 0)
    text_color_daughter_son = (9,35,138)

    # orignal font fonts/AlexBrush-Regular.ttf 
    font_name = ImageFont.truetype("fonts/SnellRoundhand/SnellBT-Bold.otf", 125, encoding="unic")
    # font2 = ImageFont.truetype("fonts/AlexBrush-Regular.ttf", 250, encoding="unic")
    font_daughter_son = ImageFont.truetype("fonts/AvenirNext/AvenirNextLTPro-Regular.otf", 70, encoding="unic")
    font_p_no = ImageFont.truetype("fonts/AvenirNext/AvenirNextLTPro-Bold.otf", 60, encoding="unic")

    # draw student's name
    d.text(location_student_name, i.title(), fill=text_color, font=font_name)

    # draw female/male 
    if sex.title() == "Male":
        d.text(location_daughter_son, "Son of", fill=text_color_daughter_son, font=font_daughter_son)
    else: 
        d.text(location_daughter_son, "Daughter of", fill=text_color_daughter_son, font=font_daughter_son)
    
    # draw father's name 
    d.text(location_father_name, father.title(), fill=text_color, font=font_name)

    # draw Phlon No. under student photo
    d.text(location_p_no, "P-No.%s" % id, fill=text_color, font=font_p_no)

    # img1.show()
    img1.save("images/pbm-2022/certificates/pbm_2022_b1_%s.jpg" % (id))
    print("(%d/%d) Certificate Created for:  %s" % (idx+1, max_no, i.title()))
    
    
print("""\n*************************
All Certificates Created! \nPlease follow me on Github for more awesome codes\n\tGitHub: https://github.com/mursalfk
*************************
""")
#Read readme.md for further instructions