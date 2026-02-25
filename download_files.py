# -*- coding: utf-8 -*-
"""
Created on Sun Oct 13 15:37:08 2019

@author: hewi
"""

#### IF error : "ModuleNotFOundError: no module named PyPDF2"
   # then uncomment line below (i.e. remove the #):
       
#pip install PyPDF2

import pandas as pd
from openpyxl import load_workbook
import PyPDF2
from pathlib import Path
import shutil, os
import os.path
import urllib
import urllib.request as req
import glob


###!!NB!! column with URL's should be called: "Pdf_URL" and the year should be in column named: "Pub_Year"

### File names will be the ID from the ID column (e.g. BR2005.pdf)

########## EDIT HERE:
    
### specify path to file containing the URLs
folder_pth = r'C:\Users\LineWienke\Documents\Python_PDF_downloader/'
list_file = 'GRI_2017_2020.xlsx'
list_pth = folder_pth + list_file

###specify Output folder 
pth = folder_pth + 'Downloaded-pdfs/'

###Specify path for existing downloads
dwn_pth = pth + "dwn/"

###Specify file for MetaData in output folder
MD = "Metadata2006_2016.xlsx"
MD_pth = pth + MD

###specify the ID column name
ID = "BRnum"

##########

### check for files already downloaded
dwn_files = glob.glob(os.path.join(dwn_pth, "*.pdf")) 
exist = [os.path.basename(f)[:-4] for f in dwn_files]

#print(exist)


### read in file for links
df = pd.read_excel(list_pth, sheet_name=0, index_col=ID)

### filter out rows with no URL in first coloum
non_empty = df.Pdf_URL.notnull() == True 
### filter out rows with no URL in secound coloum and only keep the ones that have
non_empty_2 = df.Report_Html_Address.notnull() == True
non_empty_2 = non_empty_2.loc[non_empty_2]
non_empty.update(non_empty_2)

###filter out the ones that are missing
df = df[non_empty]
df2 = df.copy()


### TO BE IMPLEMENTET
#  find which cases have alreay files in the excel sheet
#dfmd = pd.read_excel(MD_pth, sheet_name="Metadata2006_2016", index_col=ID)
#dfmd = dfmd['pdf_downloaded']
#dwn_yes = dfmd.pdf_downloaded == "yes"
#print(dwn_yes)
#dfmd = dfmd[dwn_yes]
#print(dfmd)
#df2 = df2.merge(dfmd['pdf_downloaded'], how="inner", on=ID)

#print(df2)



### filter out rows that have been downloaded
df2 = df2[~df2.index.isin(exist)]

#print(df2)

def try_to_download(Dataframe, index, column, file):
    try:
        req.urlretrieve(Dataframe.at[index, column], file)
        print("got no: " + j + "with column: " + column)
        if os.path.isfile(file):
            
            try:
                pdfFileObj = open(file, 'rb')
                #creating a pdf reader object
                pdfReader = PyPDF2.PdfReader(pdfFileObj)
                with open(file, 'rb') as pdfFileObj:
                    pdfReader = PyPDF2.PdfReader(pdfFileObj)
                    #check if there is pages in the document
                    if len(pdfReader.pages) > 0:
                        Dataframe.at[index, 'pdf_downloaded'] = "yes"
                        return True
                    else:
                        Dataframe.at[index, 'pdf_downloaded'] = "file_error"
               
            except Exception as e:
                    Dataframe.at[index, 'pdf_downloaded'] = str(e)
                    print(str(str(index)+" " + str(e)))
                 
        else:
            Dataframe.at[index, 'pdf_downloaded'] = "404"
            print("not a file")
            
    except (urllib.error.HTTPError, urllib.error.URLError, ConnectionResetError, Exception ) as e:
                Dataframe.at[index, "pdf_downloaded"] = "ERROR"
                Dataframe.at[index,"error"] = str(e)
                #print(e)
    
    return False



df2 = df2.head(5)
### loop through dataset, try to download file.
for j in df2.index:
   
   #Create the place for where to store the file and the name
    savefile = str(dwn_pth + str(j) + '.pdf')

    #try and download it with the first link
    res = try_to_download(df2, j, 'Pdf_URL', savefile) 

    #if first link dosn't work, try the second
    if(res == False):
        try_to_download(df2, j, '', savefile)

    
  
#open the excel sheet
df_existing = pd.read_excel(MD_pth, index_col=ID)

#add the pdf status from the other dataframe to the dataframe from the excel sheet, and into a new dataframe
df_combined = pd.concat([df_existing, df2.pdf_downloaded])

#overwrite the excel sheet with the previous data along with the new pdf status
df_combined.to_excel(MD_pth, index=True)



