# Certificates-Generator

We at ACM VNR VJIET, have automated the process of generating certificates of people who have attended the workshops conducted by team ACM.  
This program automatically generates the certificates of the members by giving their details as input and saves them in both .jpg and .pdf file formats for better compatibility. 
It saves a lot of time by generating certificates of multiple people at the same time within a few seconds, which otherwise takes a lot of time if to be done manually.


Introduction 
Certificates Generator is the python project to generate certificates. It also serves as GUI based application.
Technologies

Project is created with:
•	Python programming language
•	Tkinter package

Setup
To run this project, follow the below instructions:

1.	Copy the code from certificate.py python file to python IDLE or any interactive interpreter.
2.	Make sure all the related files such as sample template certificate.png and true type font file lora-bold.ttf are in the same folder.
3.	Make sure that the excel file sample certificates.xlsx, used for the details of the participants, exists on the system.
4.	Run the program.
5.	See that all required packages, modules, and libraries are installed and imported successfully.
6.	Successful generation of certificate for each member mentioned in the excel file is determined 
    when the respective row number along with name is printed on the output terminal.
7.	Finally, a Done message will be printed on successfully execution of the program.
8.	All_Certificates folder will be created in the specified directory path. Subfolders named Images and PDFs are created, 
    which stores images and portable document formats(pdf) of all generated certificates respectively.

Features
Certificates are generated and saved in both image(.png) and document(.pdf) formats.
