import os.path
import win32com.client
import time
from PyPDF2 import PdfFileWriter, PdfFileReader


def convertRTFToPDF(inputFiles,destFolderLoc,emitter=None):
    """
    Converts One or More selected RTF file into Pdf, into selected destination location.
    :param fileDict: Contains dictionary of source file location, and destination file location
        fileDict={
            'c:\\user\\jhon\\src\\file1.rtf' : 'c:\\user\\jhon\\dest',
            'c:\\user\\jhon\\src\\file2.rtf' : 'c:\\user\\jhon\\dest',
            'c:\\user\\jhon\\src\\file3.rtf' : None,
            ,
        }
    :return: True if conversion completed successfully
             False if got Error
    """
    # success = 0
    completeFiles = []
    incompleteFiles = []
    if len(inputFiles)==0:
        emitter.processSignal.emit(100)
        return completeFiles,incompleteFiles

    wdFormatPDF = 17

    # Opening Word Application
    try:
        word = win32com.client.Dispatch('Word.Application')
    except Exception as e:
        print("Warning!", "Word Application is required")
        print(e)
        return completeFiles,incompleteFiles

    emitter.processSignal.emit(0)
    for i,srcFileLoc in enumerate(inputFiles):

        inFile = os.path.basename(srcFileLoc)
        if emitter is not None:
            emitter.fileStarted.emit(inFile)# input file name
        outFile = ".".join(inFile.split('.')[:-1])
        # if dest is None, files is saved to src directory
        # destFolderLoc = os.path.dirname(srcFileLoc) if not destFolderLoc else destFolderLoc
        outFilePath = os.path.join(destFolderLoc, outFile + '.pdf')

        # if file with same name exists in dest folder
        try:
            if os.path.exists(outFilePath):
                word.Quit()
                raise FileExistsError("File already exist in dest location ," + outFilePath)
                incompleteFiles.append(os.path.basename(outFilePath))
            # Opening Doc in Word
            doc = word.Documents.Open(srcFileLoc,Visible=False)
            if not doc:
                raise  FileNotFoundError("File does not exist in location ," + srcFileLoc)
                incompleteFiles.append(os.path.basename(outFilePath))
            doc.SaveAs(outFilePath, FileFormat=wdFormatPDF)
            doc.Close()
            print('File Saved to :', outFilePath)
            completeFiles.append(os.path.basename(outFilePath))
            # success+=1
        except Exception as e:
            print('Error Occured while converting rtf to pdf')
            print(e)
            incompleteFiles.append(os.path.basename(outFilePath))

        i = i + 1
        if emitter is not None:
            emitter.fileCompleted.emit(outFile)
            emitter.processSignal.emit((i/len(inputFiles))*100)

    # completed conversion successfully
    try:
        word.Quit()
    except:
        pass
    return completeFiles,incompleteFiles

def checkFileExists(paths):
    nonExistentFiles = []
    for path in paths:
        if not os.path.exists(path):
            nonExistentFiles.append(path)

    return nonExistentFiles

def merge(paths, output, bookmark_dicts=None, toc_page_size=0):
    """
    Merge two or more pdf files into a single pdf file
    :param paths: list of source pdf files absolute location
    :param output: absolute path to output file, which is a merged pdf file
    :return:
    """
    # create a write stream
    pdfWriter = PdfFileWriter()
    # check if input file path exists
    nonExistingFiles = checkFileExists(paths)
    if len(nonExistingFiles) > 0:
        print("Following files were not found")
        print(nonExistingFiles)
        raise FileNotFoundError('Unable to find some input files, for merging operation')
    # check if resulting output pdf with same name already exists in that location
    if os.path.exists(output):
        os.remove(output)
        # raise FileExistsError('File with same name already exists, in location', output)
    # Else perform merge operation
    # read each file
    for path in paths:
        pdfReader = PdfFileReader(path)
        for page in range(pdfReader.getNumPages()):  # for all page add it to write stream
            pdfWriter.addPage(pdfReader.getPage(page))
    if bookmark_dicts is not None:
        # add bookmark
        bookmarkAdder(bookmark_dicts, pdfWriter, toc_page_size)
    # writing out to single pdf
    with open(output, mode='wb') as out:
        pdfWriter.write(out)

# bookmark adder to like navigation pane in pdf viewer which provides clickable link interface
def bookmarkAdder(bookmark_dict, pdfWriterRef, toc_page_size=0):
    """
    Adds bookmark(outlines) to the pdfWriter stream provided

    :param bookmark_dict: dictionary mapping title with page number
    :param pdfWriterRef: stream of pdf file writer
    :param parentRef: parent bookmark title for the nested child bookmark
    :return: stream of pdf file writer
    """
    for header, pgNum in bookmark_dict.items():
        pdfWriterRef.addBookmark(header, pagenum=pgNum+toc_page_size-1)

    return pdfWriterRef


# Toc generations
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch, cm

from reportlab.platypus import SimpleDocTemplate, Paragraph, Frame, KeepTogether, BaseDocTemplate, PageTemplate
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Table
from reportlab.platypus import TableStyle
from reportlab.lib import colors


def genTOC(outFilePath, bookmark_dict, runFlag=2, offset=0):
    """
    Generates Table of Contents in pdf file format with A4 page size
    :param outFilePath: output pdf file name location
    :param bookmark_dict: dict containing article title and pageNum
        {'Introduction': 5,
        'Background': 8}
    :return: outFilePath if successfully generated, else throws exception
    """
    # runFlag                        # flag to determine is it executing at first time
    # can write into if file already exists with same filename in same path
    # Create a frame
    # A4 = (210 * mm, 297 * mm)
    text_frame = Frame(
        x1=3.00 * cm,  # From left
        y1=1.5 * cm,  # From bottom
        height=26.70 * cm,
        width=18.00 * cm,
        leftPadding=0 * cm,
        bottomPadding=0 * cm,
        rightPadding=1.7 * cm,
        topPadding=0 * cm,
        showBoundary=0,
        id='text_frame')
    # Add Table
    titleTable = Table([
       ["Table of Contents"],
    ], [520])
    contentItems = []
    mainTable = [
        titleTable,
    ]
    # Add Style
    '''
    # List available fonts
    from reportlab.pdfgen import canvas
    for font in canvas.Canvas('abc').getAvailableFonts():
        print(font)
    '''
    titleTableStyle = TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTSIZE', (0, 0), (-1, -1), 18),
        ('FONTNAME', (0, 0), (-1, -1), 'Times-Bold'),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 15),
        ('LINEBELOW', (0, 0), (-1, -1), 1, colors.black),
    ])
    titleTable.setStyle(titleTableStyle)
    contentParagraphStyle = ParagraphStyle(fontName='Times-Roman', fontSize=16, name='TOCHeading2', wordWrap='LTR',
                                           leading=14, )
    # Iterate over bookmark dict and add it to table row
    for title, pageNum in bookmark_dict.items():
        contentItems.append([Paragraph(title, style=contentParagraphStyle), pageNum+offset])
    if contentItems:
        rowItemTable=(Table(contentItems, colWidths=(16.0*cm, 2*cm)))
        mainTable.append(rowItemTable)
    # mainTable = KeepInFrame(0,0, mainTable, mode='shrink',)
    mainTable=KeepTogether(mainTable)
    # Create story
    story = [mainTable]
    # story.append(mainTable)
    pdf = BaseDocTemplate(outFilePath, pagesize=A4,)
    # page template
    pageTemplate = PageTemplate(id='pageTemplate',
                                frames=[text_frame]
                                )
    pdf.addPageTemplates(pageTemplate)
    pdf.build(story)
    print("toc page size", pdf.page)
    # creating pdf first to determine file total page number in toc
    if runFlag == 2:
        # delete existing created toc file
        print('Deleting Created toc file')
        os.remove(outFilePath)
        genTOC(outFilePath, bookmark_dict, runFlag=runFlag-1, offset=pdf.page)
    return pdf.page, outFilePath

def rowContentItem(title, pageNum):
    """
    Generates Row Table Dynamically for each header title and pageNum which is latter inserted as row to main table.
    :param title:
    :param pageNum:
    :return: Table
    """
    # Global Style for row content and paragraph
    contentTableStyle = TableStyle([
        ('FONTSIZE', (0, 0), (-1, -1), 14),
        ('FONTNAME', (0, 0), (-1, -1), 'Times-Roman'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ])

    return Table([
        [title, pageNum],
    ], [500, 40], style=contentTableStyle)

# Bookmark dictionary generation
import fitz
import re

def createBookmarkDict(xfile, xStringList):
    """
    Extracts page Number for headers with bold text, in pdf and match with provided headerList and creates bookmarks

    :param xfile: Input file path to be read
    :param xStringList: list of headers whose pagenum location is to be find. like ['Introduction', 'Background']
    :return: bookmark dict
        {'Introduction': 5,
        'Background': 8}
    """
    pdfDoc = fitz.open(xfile)  # open existing pdf
    pageFound = -1
    headerPageLocDict = {}

    for page in pdfDoc:
        bolderTexts = []  # bold text in that page
        blocks = page.get_text('dict', flags=11)['blocks']
        for b in blocks:
            for l in b["lines"]:  # for every line
                for s in l["spans"]:  # iterate over every span
                    if s["flags"] in [9,20,16]:  # 20 is bold text say 20 or 16 is bold for header element
                        bolderTexts.append(s["text"])
        bolderTexts = " ".join(bolderTexts)

        # check for every text in headerList and matching with boldText in that page
        for xString in xStringList:
            if not headerPageLocDict.get(xString):  # if pageNum for xString is not set/found
                # result = re.search(xString.replace('\n','  '),bolderTexts)  # Used .replace() since, boldText is usually like for e.g "Using Regular Expressions" is converted to "UsingRegularExpressions"
                # xStringList.remove(xString)
                if xString.replace(' \n','  ').replace('\n', '  ') in bolderTexts.replace(' \n','  ').replace('\n', '  '):
                    # print(xString)
                    # found
                    headerPageLocDict[xString] = page.number + 1  # page starts with 0 while reading programmatically
                    # xStringList.remove(xString)
                    # break
                # else:
                #     headerPageLocDict[xString] = None

    print(headerPageLocDict)
    return headerPageLocDict


#
if __name__ == "__main__":

    cwd =  r'C:\Users\dipeshs\pdf\merged'

    # Testing rtf to pdf conversion
    fileDict = {
        r'C:\Users\dipeshs\pdf\Listing 16.1.rtf': cwd,
        r'C:\Users\dipeshs\pdf\Listing 16.2.rtf': None,
        r'C:\Users\dipeshs\pdf\Listing 16.3.rtf':cwd,
    }

    if convertRTFToPDF(fileDict,r'C:\Users\dipeshs\pdf\merged'):
        print("RTF to PDF conversion Completed Successfully")
    else:
        print("Error occured while conversion of rtf to pdf")

    # testing merging operation

    paths = [os.path.join(r'C:\Users\dipeshs\pdf\merged', 'Listing 16.1.pdf'),
             os.path.join(r'C:\Users\dipeshs\pdf\merged', 'Listing 16.2.pdf'),
             os.path.join(r'C:\Users\dipeshs\pdf\merged', 'Listing 16.3.pdf'), ]

    mergedOutPath = os.path.join(cwd, 'mergedfile.pdf')
    merge(paths, mergedOutPath)

    # Test : Getting bookmark dictionary created
    headerList = [
        "16.1",
        "16.2",
        "16.3",
        "16.4"
    ]
    bookmarkDict = createBookmarkDict(mergedOutPath, headerList)

    # Test : Creating TOC file
    tocFilePath = os.path.join(cwd, 'tocccc.pdf')
    tocfile = genTOC(tocFilePath, bookmarkDict)

    # Merging TOC file to previous merged file
    paths = [tocFilePath,mergedOutPath]
    outPath = os.path.join(cwd, 'final.pdf')
    merge(paths, outPath, bookmarkDict)

    import  fitz

    pdfFile = fitz.open(outPath)

