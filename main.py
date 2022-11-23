import datetime
import os
import sys
import pandas as pd
from pathlib import Path
import shutil
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

files = []

# Clear and create a new `outputs` folder
if os.path.exists('outputs'):
    shutil.rmtree('outputs')
os.makedirs('outputs')

# Loop through `templates` folder to get the templates
for dirpath, dirnames, filenames in os.walk('templates'):
    for filename in filenames:
        # Ignore any file that is not a .docx
        if not filename.endswith('.docx'):
            continue
        file = os.path.join(dirpath, filename)
        files.append(file)

# Load data workbook
for f in os.listdir('inputs'):
    if f.endswith('xlsx'):
        excel_file = f
        break
data_df = pd.read_excel('inputs/' + f, sheet_name=0, engine='openpyxl')
data_df.dropna(axis=1, how='all', inplace=True)

# Get variables from `data` workbook
variables = data_df.columns.to_list()

# Loop through each template file
for file in files:

    # Open the Word document as a template
    template = DocxTemplate(file)

    # Render template for each row in `data` workbook
    for row in data_df.itertuples():

        # Create a dictionary to store the variables and their values
        context = {}
        # Add the entire row as values
        for variable in variables:
            variable_value = getattr(row, variable)
            if isinstance(variable_value, datetime.date):
                try:
                    context[variable] = variable_value.strftime('%d/%m/%Y')
                except ValueError:
                    continue
            else:
                context[variable] = variable_value  

        # Add images as variables
        images = []
        for dirpath, dirnames, filenames in os.walk('inputs'):
            for filename in filenames:
                # Ignore any file that is not a picture
                if not os.path.splitext(filename)[-1] in ['.jpg', '.png']:
                    continue
                img_file = os.path.join(dirpath, filename)
                images.append((Path(filename).stem, img_file))
        
        for img_name, img_path in images:
            context[img_name] = InlineImage(template, img_path, height=Mm(10))

        # Render template
        template.render(context)
       
        # Save document using first column of data workbook as identifier
        file_name_without_extension = Path(file).stem
        doc_name = f'{file_name_without_extension}-{context[variables[0]]}.docx'
        template.save('outputs/' + doc_name)
        print('Creating: ' + doc_name )

print('Documents successfully rendered.')