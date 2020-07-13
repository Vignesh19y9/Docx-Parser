import pypandoc
import mammoth
import shutil            
from bs4 import BeautifulSoup
import re
import os
import PIL.Image as Image
import csv


file_name="input/4table.docx"         # Source Docx file name
csv_filename="output/result.csv"       # Destination Csv file name
file='old'                      # Specify csv file should be new or old
mediaPath="images"              #Image directory


# =============================================================================
# Creat csv file
# =============================================================================
field = ['Question', 'Type', 'Option1', 'Option2','Option3','Option4','Solution','Answer','Marks'] 

def CreateFile(csv_filename):  
    global Created_file
    
    if  not os.path.exists(csv_filename):                       
        print("Created new file >",csv_filename)        
        with open(csv_filename, 'w') as csvfile: 
            csvwriter = csv.writer(csvfile) 
            csvwriter.writerow(field)
  
    else:
        print("File Already Exist >",csv_filename)
        i=0
        while True:
            newName=csv_filename.split('.csv')[0]+str(i)
            newName=newName+".csv"
            
            if  os.path.exists(newName):
                i=i+1
            else:
                print("Created new file >",newName)
                Created_file=newName

                with open(newName, 'w') as csvfile: 
                    csvwriter = csv.writer(csvfile) 
                    csvwriter.writerow(field)
               
                break
    
#------------------------------------------------------   
Created_file=csv_filename

                
if file=='new':
    
    CreateFile(csv_filename)
    
elif file=='old':
    #checks old file
    if  not os.path.exists(csv_filename):
        print("File Doesnt Exist >",csv_filename)
        CreateFile(csv_filename)
        
    else:
        #checks if old file has fields 

        newfields=[]
        try:            
            with open(csv_filename, 'r',encoding='utf-8') as csvfile: 
                # creating a csv reader object 
                csvreader = csv.reader(csvfile)                   
                # extracting field names through first row will be not preset if csv file nave nothing 
                oldfields = next(csvreader)
                
                if len(oldfields)!=len(field):
                    print("Fields mismatch creating a new file")
                    CreateFile(csv_filename)
                    
        except:
            print("Empty file")
            with open(csv_filename, 'w') as csvfile: 
                csvwriter = csv.writer(csvfile) 
                csvwriter.writerow(field)

                
        
# =============================================================================
# Finding total no of tables in file
# =============================================================================

def FindTable(html):
    
    parsed_html = BeautifulSoup(html,"lxml")            # Convert to BeautifulSoup element

    Total_table=parsed_html.find_all('table')           # Checks data contain <table> </table> if yes returns data in b/w
    print("Total table found =",len(Total_table))
    
    InnerTable=[]
    
    for tno,table in enumerate(Total_table):    # Checks all the detected table if it is inner or outer
                                                # If inner table is found it is removed from outer table list
        if table.find_all('table'):
            inner=table.find_all('table')
            
            InnerTable.append(len(inner))
            
            for i in inner:     
                if i in Total_table:
                    
                    
                    Total_table.remove(i)
                    print("Inner table found in table {} removing it from Total tables count".format(tno))
                    print("Total table found =",len(Total_table))
                    
        else:InnerTable.append(0)            
#    print("inner table count =",InnerTable) 
    return Total_table
# =============================================================================
# clean
# =============================================================================
tags="<td>|</td>|<tr>|</tr>"
#valid_images = [".jpg",".jpeg",".gif",".png",".tga",".pgm",".tiff"]
    
def clean(cell):                                # Function to clean tags from the data
    cell=re.sub(tags,"", cell)
    return cell

# =============================================================================
# SaveImage
# =============================================================================
mediaHist={}                                    # For storing new image name since we are using 2 html file

def SaveImage(cell):
#    print(cell)
    
    if not os.path.exists(mediaPath):           # Checking the existence of Media path 
        os.makedirs(mediaPath)
        
    filelist=os.listdir(mediaPath) 
             # Listing all files in Media path    
    imgList=re.findall('<img(?![^>]*\balt=)[^>]*?>',cell)      # Finiding all <Img> tags in cell
#    print(imgList)
    ImgNewPath=''
    for imgPathFull in imgList:                     # Iterating over Img tags
        
        imgPathFormated=imgPathFull.replace('\\','/')
        
        imgPath=re.findall('src="(.*?)"',imgPathFormated)[0] # Finding source tag of image
        imgName=imgPath.split('/')[-1]

        
        if imgPath in mediaHist:
            ImgNewPath=mediaHist[imgPath]
                       
        #--------------------------------------------------    
        elif imgName not in filelist:           # If Img name is not in Media path simply save it
            ImgNewPath=os.path.join(mediaPath,imgName)
        else:
            
            imgNamePart=imgName.split('.')[0]   # Splitting Img name and 
            ext=imgName.split(imgNamePart)[-1]  # its extension            
            i=1                                 
            while True:
                newName=imgNamePart+str(i)+ext  # Creates new name 
                if newName in filelist:         # Checking if the new name is also in the Media path
                    i=i+1
                    continue
                else:
                    break
                        
            ImgNewPath=os.path.join(mediaPath,newName)
        #--------------------------------------------------------   
        img=Image.open(imgPath)                 # Opens image and save it in new directory
        img.save(ImgNewPath)                    
        
        mediaHist[imgPath] =ImgNewPath
        ImgNewPathFormatted="<img src= "+ImgNewPath+" >"    # Adding source tags
        cell=cell.replace(imgPathFull,ImgNewPathFormatted)
        
    return cell
    
    
# =============================================================================
# Parse Table    
# =============================================================================
def ParseTable(Total_table):
    
    table_data=[]                                   # For storing table values after process
    for tno,table in enumerate(Total_table):
        
        table=str(table)                            # Stringfy
        rows= re.split('\n',table)                  # Splitting with /n to form each tag as a list
#        print(rows)
        rows.pop(0)                                 # Deleting <table> top
        rows.pop(-1)                                # Deleting </table> end 
        
        rowData=[]                                  # For storing row values after process    
        cellData=[]                                 # For storing cell values after process                                
        cellTable=0                                 # Count for inner tables if any
        miniTable=''
        
        for cell in rows: 
            if re.search('src="(.*?)"',cell):       # Iterate over each cell 
                    cell=SaveImage(cell) 
                    
            if re.findall('<tr>',cell):             # Checks if it has a row starting if not skips 
                #---------------------------------------------------        
                if cellTable==0:                    # Checks the row starting is not a row of innner tables                 
                    cell=clean(cell) 
                                
                    if len(cellData)!=0:            # If detected <tr is a new row and the previous row data is in CellData
                        rowData.append(cellData)    # We move it to rowdata and
                        
                        cellData=[]                 # Clears the data and
                        cellData.append(cell)
                                                                 
                    else:cellData.append(cell)      # Adds new data  
                        
                else:                
                    miniTable=miniTable+cell           # Table data need not be cleaned if table is present
                #---------------------------------------------------  
            if re.findall('<table>',cell):          # Checking for starting inner table            
                cellTable=cellTable+1               # If found the count will be incremented
                miniTable=miniTable+cell            # inserting <table> tag
                
            elif re.findall('</table>',cell):         # Checking for inner table end       
                cellTable=cellTable-1               # Decrementing the table count
                cell=clean(cell)                    # </table> might contain other tag so we need to clean it
                miniTable=miniTable+cell
                cellData.append(miniTable)
                miniTable=[]
                
                
            #----------------------------------------------------------
            # All data inside table is searched and added                      
            elif re.findall('<td>',cell) or re.findall('</td>',cell) or re.findall('</tr>',cell):            
    
                if cellTable==0:                   # If not a inner table data we need to clean it
                    cell=clean(cell)
                    cellData.append(cell)
                else:miniTable=miniTable+cell
                                                                         
            else:continue
                                                   # deleting '' cells
            cellData=list(filter(('').__ne__, cellData))
        
        rowData.append(cellData)                   # The final tag <tr to check in abouve condition so add last row a end of a table   
        table_data.append(rowData)                 # combining all row to table
    return table_data        
# =============================================================================
# Html cleaning
# =============================================================================

def cleanHtml(rhtml):
                                            # Removes all unwanted tags and adds new line if necessary
    rhtml=re.sub('<p>|</p>','',rhtml)
    rhtml=re.sub('<thead>','',rhtml)
    rhtml=re.sub('</thead>','',rhtml)
    rhtml=re.sub('<tbody>','',rhtml)
    rhtml=re.sub('</tbody>','',rhtml)
    rhtml=re.sub('\r','',rhtml)
#    
#    
    rhtml=re.sub('<tr(.*?)>','<tr>',rhtml)
    rhtml=re.sub('<th>','<td>',rhtml)
    rhtml=re.sub('</th>','</td>',rhtml)
    rhtml=re.sub('<th.*?>','<td>',rhtml)
    rhtml=re.sub('<td.*?>','<td>',rhtml)
    rhtml=re.sub('<td>','<td>',rhtml)
#    
    rhtml=re.sub('u2004','',rhtml)

    
    if len(re.findall('\n',rhtml))==0:
        
        rhtml=re.sub('<table>','\n<table>\n',rhtml)
        rhtml=re.sub('</table>','</table>\n',rhtml)
        rhtml=re.sub('<tr>','<tr>\n',rhtml)
        rhtml=re.sub('</td>','</td>\n',rhtml)
        rhtml=re.sub('</tr>','</tr>\n',rhtml)
    
    rhtml=re.sub('\r','',rhtml)    
#    rhtml=re.sub('\n\n','\n',rhtml)
    return rhtml

# =============================================================================
# Html Mammoth helper
# =============================================================================
class ImageWriter(object):
    def __init__(self, output_dir):
        self._output_dir = output_dir
        self._image_number = 1

    def __call__(self, element):
        default_name=element.content_type.partition("/")[0]
        extension = element.content_type.partition("/")[2]
        image_filename = "{0}.{1}".format(default_name+str(self._image_number), extension)
        
        with open(os.path.join(self._output_dir, image_filename), "wb") as image_dest:
            with element.open() as image_source:
                shutil.copyfileobj(image_source, image_dest)

        self._image_number += 1

        return {"src": os.path.join(self._output_dir, image_filename)}
    
# =============================================================================
# Raw html    
# =============================================================================
# Conversions of docx to html by pypandoc
rhtml2 = pypandoc.convert_file(file_name, 'html',extra_args=['--extract-media=temp'])#outputfile="3table.html"
   
convert_image = mammoth.images.inline(ImageWriter('temp/media'))
rhtml1 = mammoth.convert_to_html(file_name,convert_image=convert_image).value # Conversions of docx to html by Mammoth

html1=cleanHtml(rhtml1)             # Cleans the data and formatts it
html2=cleanHtml(rhtml2)

tableDet1=FindTable(html1)          # Detects and returns the tables
tableDet2=FindTable(html2)

tableData1=ParseTable(tableDet1)    # Extract data, equ, image path from tables
tableData2=ParseTable(tableDet2)


finalTable=[]
                                    # combines 2 forms into 1 and store it as dictionary
for t1 ,t2 in zip(tableData1,tableData2):
    dataDict={}
    
    optNo=1
    for r1,r2 in zip(t1,t2):
        
        if r2[0]=='Option':
            r2[0]='Option'+str(optNo)
            optNo=optNo+1
                        
        if len(r1)>len(r2):       
            r2.append(r1[-1])
            
        dataDict[r2[0]]=','.join(r2[1:])
        
    finalTable.append(dataDict)
    

# Writing file as csv 
with open(Created_file, 'a', newline='',encoding='utf-8') as csvfile:
    writer = csv.DictWriter(csvfile, fieldnames=field)

    writer.writerows(finalTable)

shutil.rmtree('temp') # Removes temporary file
        



