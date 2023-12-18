#Compiles a PDF by merging together Bay View's sidewalk coversheet report with their
#individual repair report(s) for each Lot.
#
#Resource: This website was relied on heavily to understand the PyPDF2 function:
#https://caendkoelsch.wordpress.com/2019/05/10/merging-multiple-pdfs-into-a-single-pdf/
#Resource: This website helped me sort the Form folder directly alphanumerically:
#https://stackoverflow.com/questions/2669059/how-to-sort-alpha-numeric-set-in-python
#
#Script was created on 2-25-22 by Paul McCord


#import modules
import os
import PyPDF2
import re

#Function sorts file names alphanumerically.
#This function was needed as the files in the directories had too be sorted appropriately
#for the final output pages to be sorted correctly.
#The problem was, for example, Lot 12 would be placed in front of Lot 2 in the unordered
#directory. This function corrects that (ie, it puts Lot 2 in front of Lot 12, for example)
def sorted_nicely(l):
    convert = lambda text: int(text) if text.isdigit() else text 
    alphanum_key = lambda key: [ convert(c) for c in re.split('([0-9]+)', key) ] 
    return sorted(l, key = alphanum_key)

#This function does the bulk of the PDF compilation.
#Each Lot (ie, coversheet) map is iterated over and if matching individual repair
#maps are found, then those are added immediately after the coversheet map in the PDF
def compile_pdf(pdf_writer, coversheet_path, repair_sheet_path, path_lot_pdf_name, search_lot_pdf_name, repair_list):
    print(path_lot_pdf_name)
    #Create file name to match name in the Lot directory
    lot_pdf_ind_name = "Lot_Map_" + path_lot_pdf_name + ".pdf"
    #Open Lot PDF file and read to PDFReader
    pdf1File = open(os.path.join(coversheet_path, lot_pdf_ind_name), 'rb')
    pdf1Reader = PyPDF2.PdfFileReader(pdf1File)

    #Iterate over all pages of Lot PDF (there's only one) and add page to pdf_writer
    for pageNum in range(pdf1Reader.numPages):
        pageObj = pdf1Reader.getPage(pageNum)
        pdf_writer.addPage(pageObj)
    #pdf1File.close()
    
    #Count occurences of Lot name in the list of Repair files
    #This is needed to add the letter suffix to the Repair file so as to match it
    #in the directory
    repair_occ_counter = 0
    for r in repair_list:
        if r == search_lot_pdf_name:
            repair_occ_counter += 1

    #Repair dictionary to match occurrence to letter suffix
    repair_pdf_dict = {
        0: 'DoesNotExist',
        1: 'A',
        2: 'B',
        3: 'C',
        4: 'D',
        5: 'E',
        6: 'F',
        7: 'G',
        8: 'H',
        9: 'I',
        10: 'J',
        11: 'K',
        12: 'L',
        13: 'M',
        14: 'N',
        15: 'O',
        16: 'P',
        17: 'Q',
        18: 'R'
    }
    
    
    for x in range(repair_occ_counter):
        #Create file name to match name in the Repair directory
        repair_pdf_ind_name = "Block_and_Lot_" + search_lot_pdf_name + "-" + repair_pdf_dict[x + 1] + ".pdf"
        print(repair_pdf_ind_name)
        #Open Repair PDF file and read to PDFReader
        pdf2File = open(os.path.join(repair_sheet_path, repair_pdf_ind_name), 'rb')
        pdf2Reader = PyPDF2.PdfFileReader(pdf2File)

        #Iterate over all pages of Repair PDF (there's only one) and add page to pdf_writer
        for pageNum in range(pdf2Reader.numPages):
            pageObj = pdf2Reader.getPage(pageNum)
            pdf_writer.addPage(pageObj)
        #pdf2File.close()
        print(f"added {repair_pdf_ind_name} to PDF")

    return pdf1File, pdf2File


if __name__ == "__main__":
    coversheet_path = r"P:\7700_7799\7740210010_Bay_View_GIS_General_Services\_GIS\ArcLayouts\Sidewalk_Repair_Lots_20220223"
    repair_sheet_path = r"P:\7700_7799\7740210010_Bay_View_GIS_General_Services\_GIS\ArcLayouts\Sidewalk_Repair_IndividualRepairs_20220223"

    #Create a new PDF Writer object
    #This object will be used to write the pdfs to an output document
    pdf_writer = PyPDF2.PdfFileWriter()

    #get file names from coversheet folder
    sorted_list = []
    list_dir = os.listdir(coversheet_path)

    #Run sorted_nicely function to put Lot files in the correct order
    s = set(list_dir)
    for x in sorted_nicely(s):
        sorted_list.append(x)
    
    for f in sorted_list:
        if os.path.isfile(os.path.join(coversheet_path, f)):
            #coversheet_lst.append(f)
            pdf_name = f
            #Process to format Lot names and strip out unnecessary text
            sub_str = pdf_name.split("_", 2)[-1]
            path_lot_pdf_name = sub_str.split(".")[0] #the name of the file in file explorer uses an underscore between Block and Lot
            search_lot_pdf_name = sub_str.replace('_', '-').split(".")[0] # the name that is matched to the repair PDF uses a hyphen

            #get file names from repair sheet folder
            repair_pdf_list = []
            for f in os.listdir(repair_sheet_path):
                if os.path.isfile(os.path.join(repair_sheet_path, f)):
                    pdf_name = f
                    #Process to format Repair names and strip out unnecessary text
                    sub_str = pdf_name.split("_", 3)[-1]
                    repair_pdf_name = "-".join(sub_str.split('-')[:2])
                    repair_pdf_list.append(repair_pdf_name)

            if search_lot_pdf_name in repair_pdf_list:
                #If the Lot name matches the Repair name, run the compile_pdf function
                pdf1File, pdf2File = compile_pdf(pdf_writer, coversheet_path, repair_sheet_path, path_lot_pdf_name, search_lot_pdf_name, repair_pdf_list)

    output_path = r"P:\7700_7799\7740210010_Bay_View_GIS_General_Services\_GIS\ArcLayouts\c2021_BVA_Sidewalk_Repair_Report_FINAL"
    pdfOutputFile = open(os.path.join(output_path, 'c2021_Sidewalk_Repairs_All.pdf'), 'wb')
    pdf_writer.write(pdfOutputFile)
    pdf1File.close()
    pdf2File.close()
    pdfOutputFile.close()
    


    
