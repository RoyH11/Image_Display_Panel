# Automatically reading images from DICOM files and display them in panel.

Last updated: 1/24/2025

For the current process, we assume we already have the summary csv file that including image details and dimension information. We will use the `pandas` library to read the csv file and `pydicom` library to read the DICOM files.

## Key libraries
- pandas
- pydicom
- python-pptx
- PIL
- cv2

To install the libraries, we can use pip install command. For example, to install the pandas library, we can use the following command:
```bash
pip install pandas
```

## Creating the initial pptx file 
Creating an initial pptx serves the purpose of creating a layout for the image display panel. We will use the `python-pptx` library to create the pptx file.

Referring to the `1002_get_optic_disc.ipynb` file, we will create a pptx file with the layout of the image display panel.

1. Selecting a random participant and creating a pptx file with the layout of the image display panel. For example, we will use participant 1002 in the summary csv file. 

```python
df = df[(df['participant_id'] == 1002) & (df['anatomic_region'].str.contains("optic", case=False))]
```

2. Creating a pptx file, each slide will contain one image with description and image dimension information.

```python
# create a new powerpoint presentation
from pptx import Presentation
from pptx.util import Inches
import pydicom
from pydicom.pixel_data_handlers.util import convert_color_space
from PIL import Image
import matplotlib.pyplot as plt
from io import BytesIO
import cv2
import pandas as pd
import os