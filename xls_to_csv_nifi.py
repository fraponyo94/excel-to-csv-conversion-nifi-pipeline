import sys
import xlrd 
import unicodecsv
import os
import glob
import uuid 

os.chdir(sys.argv[1])



def xlsToCsvConverter(filename):    

    # Get the file extension to filter out xls only
    filename_extension = filename.rsplit('/', 1)[-1].rsplit('.', 1)[-1]
     
    try:
        if filename_extension in ("xls","XLS"):
            # load the workbook for xls 
            work_book = xlrd.open_workbook(filename)

            # Check the number of sheets in the workbook.
            no_of_sheets = work_book.nsheets

            # Loop through the all the sheets.
            for sheet_number in range(no_of_sheets):
                try:
                    # Open the sheet by index for xls sheets
                    sheet = work_book.sheet_by_index(sheet_number)

                    # Open the csv file in binary write mode.
                    with open( output_dir+str(uuid.uuid1())+"-"+"%s.csv" %(sheet.name.replace(" ","")), "wb") as file:
                        # Uses unicodecsv, so it will handle Unicode characters.
                        csv_out = unicodecsv.writer(file, encoding='utf-8')                    

                        header = [cell.value for cell in sheet.row(0)]
                        csv_out.writerow(header)

                        # Loop through the rows of the sheet and write to csv file.
                        for row_number in range(sheet.nrows):
                            csv_out.writerow(sheet.row_values(row_number))

                        # Close the csv file.
                        file.close()                        
                        
                    
                except Exception as e:
                    raise e

        # Remove already converted file
        #os.remove(filename)
    except Exception as e:
        raise e

def OutputFileDir(input_dir):    

    source_dir = os.path.dirname(input_dir) 

    output_home = os.path.dirname(source_dir)

    output_dir = output_home+"/output/"


    # Check if output folder exist if not create
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)

    return output_dir


if __name__ == '__main__':
    # Output filedir
    output_dir = OutputFileDir(sys.argv[1])

    # select only files with extension xls,XLS,xlsx,XLSX
    files = glob.glob('*.xls')
    files.extend(glob.glob('*.XLS'))

    for file in files:  
        try:
            xlsToCsvConverter(file)
            os.remove(file)
        except Exception as e:
            raise e         
       
  

    
    


        


