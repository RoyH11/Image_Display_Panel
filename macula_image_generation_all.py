import json
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import pandas as pd
import os
from PIL import Image, ImageDraw, ImageFont
import pydicom
from pydicom.pixel_data_handlers.util import convert_color_space
import matplotlib.pyplot as plt
import cv2
from io import BytesIO

SCALER = 2


def get_all_patient_ids(image_data):
    df = pd.read_csv(image_data)
    patient_ids = df['participant_id'].unique()
    # sort the patient ids, ascending order
    patient_ids = sorted(patient_ids)
    return patient_ids




def load_data(image_data, patient_id):
    df = pd.read_csv(image_data)
    df = df[df['participant_id'] == patient_id]
    # create a dictionary
    macula_image_dict = {}

    for index, row in df.iterrows():
        key = str(
            row['manufacturer'] + "_" 
            + row['manufacturers_model_name'] 
            + "_" + row['laterality'] + "_" 
            + row['anatomic_region'] + "_" 
            + row['imaging'])
        
        image_path = r"L:\Lab\Roy\AI-READi-Roy" + row['filepath']

        ds = pydicom.dcmread(image_path)
        color_code = ds.PhotometricInterpretation
        spacing = ds.PixelSpacing
        image_size_mm = (ds.Rows * spacing[0], ds.Columns * spacing[1])
        image_size_px = (ds.Rows, ds.Columns)

        # convert the image to RGB
        if (color_code == "YBR_FULL_422"): 
            img = convert_color_space(ds.pixel_array, "YBR_FULL_422", "RGB")
        elif (color_code == "MONOCHROME1" or color_code == "MONOCHROME2"): 
            img = ds.pixel_array

        value = [img, image_size_px, image_size_mm]

        macula_image_dict[key] = value

    return macula_image_dict



def macula_image_panel(slide_info, output_image_file, image_data, patient_id): 
    # Load the image
    # image_data is a csv file with the following columns:
    # participant_id,manufacturer,manufacturers_model_name,laterality,anatomic_region,imaging,height,width,color_channel_dimension,sop_instance_uid,filepath
    # The filepath column contains the path to the image
    # get the rows for participant_id == 1001
    # load in the image data as a pandas dataframe

    image_dict = image_data

    # Get the canvas size, round to the nearest integer
    canvas_width = round(slide_info['canvas_size']['width'] * SCALER)
    canvas_height = round(slide_info['canvas_size']['height'] * SCALER)

    # Create a blank image
    img = Image.new('RGB', (canvas_width, canvas_height), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    # load a font 
    try: 
        font = ImageFont.truetype("arial.ttf", 14 * SCALER)
    except IOError:
        font = ImageFont.load_default()

    # load title font
    try: 
        title_font = ImageFont.truetype("arial.ttf", 40 * SCALER)
    except IOError:
        title_font = ImageFont.load_default()

    # load mid font
    try: 
        mid_font = ImageFont.truetype("arial.ttf", 20 * SCALER)
    except IOError:
        mid_font = ImageFont.load_default()

    # draw the image
    for element in slide_info["elements"]:
        element_type = element["type"]
        position = element["position"]
        position_x = position["x"] * SCALER
        position_y = position["y"] * SCALER

        width = element["size"]["width"] * SCALER

        # draw text 
        if element_type == "text": 
            content = element["content"]

            # edit the title
            if "1001" in content: 
                content = content.replace("1001", str(patient_id))

            # if content is empty, get the content from the image_dict            
            if content == "": 
                label = element["label"]

                if label in image_dict:
                    # get the pixel size of the image
                    # get the image size in mm
                    content_1 = image_dict[label][1]
                    content_2 = image_dict[label][2]
                    # get the elements in the tuple
                    # format the content in "AxB\nCxD  mm"
                    content = f"{round(content_1[0])}x{round(content_1[1])}\n{round(content_2[0])}x{round(content_2[1])} mm"
                else: 
                    content = None

            if content: 
                if "Participant" in content: 
                    text_position = (round(position_x + 80*SCALER), round(position_y)) # adjust the position
                    draw.text(text_position, content, font=title_font, fill="black")
                elif len(content) <= 5: 
                    text_position = (round(position_x + 5*SCALER), round(position_y))  # adjust the position
                    draw.text(text_position, content, font=mid_font, fill="black")
                else:
                    text_position = (round(position_x), round(position_y))
                    draw.text(text_position, content, font=font, fill="black")

        # draw image
        elif element_type == "image": 
            label = element["label"]
            if label in image_dict: 
                image = image_dict[label][0]
                image = Image.fromarray(image)
                height = width * (image.size[1] / image.size[0])
                image = image.resize((round(width), round(height)))
                img.paste(image, (round(position_x), round(position_y)))

        # draw line
        elif element_type == "line": 
            start = (round(position_x), round(position_y))
            end = (round(position_x + element["size"]["width"]), round(position_y + element["size"]["height"]*SCALER))
            draw.line([start, end], fill="black", width=2)


    # save the final image
    try: 
        img.save(output_image_file)
    except Exception as e: 
        print(f"Error saving the image: {e}")        




# main function-------------------------------------------------------------------------------------------------
if __name__ == "__main__": 
    json_file = r"L:\Lab\Roy\AI-READi-Roy\visualization_fundus_macula\Macula_Fundus_Layout_reordered.json"
    output_image_folder = r"L:\Lab\Roy\AI-READi-Roy\fundus demo images"
    image_data = r"L:\Lab\Roy\AI-READi-Roy\visualization_fundus_macula\All_Macula.csv"

    patient_ids = get_all_patient_ids(image_data)
    # load json file
    with open(json_file) as f: 
        slide_info = json.load(f)


    for patient_id in patient_ids[10:]:
        macula_image_dict = load_data(image_data, patient_id)
        output_image_file = os.path.join(output_image_folder, f"fundus_macula_{patient_id}.png")
        macula_image_panel(slide_info, output_image_file, macula_image_dict, patient_id)