# Fill Microsoft Word Forms
> At work, many of us have Microsoft Word forms that we need to fill on a regular basis.
> Google Colab [_here_](https://colab.research.google.com/drive/1ZAGiYGu0e-AAXBrE_TFoYv7Rkh8mhGjn) (to use this script on the cloud without any installation).


Uses cases:
- HR dept: Leave form, Claim form
- Finance dept: Purchase requisition form
- Sales dept: Customer form
- Customer service dept: Feedback form
etc.


## Table of Contents
* [General Info](#general-information)
* [Technologies Used](#technologies-used)
* [Features](#features)
* [Screenshots](#screenshots)
* [Setup](#setup)
* [Usage](#usage)
* [Project Status](#project-status)
* [To Do](#to-do)
* [Acknowledgements](#acknowledgements)
* [Contact](#contact)


## General Information
The purpose of this project is to:
- deepen my knowledge of traversing local files and folders.
- eliminate the hassle of copying information from an Excel worksheet to a Word document.


## Technologies Used
- python==3.9.5
- docxtpl==0.16.4
- pandas==1.2.4


## Features
- Can insert strings, dates, and numbers from each row of an Excel worksheet to a Word document (i.e. filling up the form)
- Can also insert images such as logos, signatures, etc.


## Screenshots


## Setup
If you don't have pandas or docxtpl, using pip to install them:
`$ pip install pandas`
`$ pip install docxtpl`

1. Download the [zipped files](https://github.com/chuawt/fill-forms/archive/refs/heads/main.zip) from the GitHub repository.
1. Unzip the files to a local directory.
1. Do not delete or rename the `inputs` and `templates` folders.
1. In the `inputs` folder, there can only be one Excel workbook. The first sheet of the workbook will be used as data inputs. 
1. Make sure that the headers are only one word (i.e. no spaces allowed).
1. Name all images as "img_<filename>" and make sure that they are all one word. 
1. The headers of the first worksheet of the workbook and the image names then becomes the stored variables.
1. In the `templates` folder, add in .docx files as templates. For each of the .docx files, use the double brackets {{ }} syntax as a placeholder, e.g. {{ Agent }}, {{ Date }}, {{ img_me }}  
1. Run the script mentioned in the below section. 
1. The output files will be in a new `outputs` folder.


## Usage
[For Mac] In Terminal, navigate to the local directory and use command:
`$ python3 main.py`
[For Windows] In command prompt, navigate to the local directory and use command:
`PS> python main.py`


# Project Status
Project is: _completed_.
Version: 1.0.0


## To Do
Some ways to extend this project:
- Refractor the code to allow users to easily modify the default settings without touching the code. 


## Acknowledgements
Many thanks to the entire Sales Department of my workplace for coming to me with their pain points for the idea of this project.


## Contact
Created by [@chuawt](https://chuawt.github.io) - feel free to contact me!