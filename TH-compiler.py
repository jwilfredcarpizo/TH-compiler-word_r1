import os
import sys
import win32com.client as win32
import datetime

welcome_text = '''
(C) 2023 AMH Philippines, Inc.
Written by: Enrico Abcede

This program compiles all SVG time histories into MS Office 365 Word .docx. Please Follow the folder structure:

~/path/to/dir/
┃-input_info.txt
/svg/
  ┃-Matched-<siteName> <RP>_<THNum> <component> <vector>.svg
  ┃-Matched-S 2500_01 H1 1-A.svg
  ┃-Matched-S 2500_01 H1 2-V.svg
  ┃-Matched-S 2500_01 H1 3-D.svg
  ┃-Matched-S 2500_02 H1 1-A.svg
  ...

The input_info.txt file must contain the following text:

<Project Name>
<Report Type>
<Annex Index>
<THNum> <TH record location> (<YYYY>)
02 Loma Prieta, California (1989)
03 Hector Mine, California (1999)
04 Kocaeli, Turkey (1999)
05 Chi-Chi, Taiwan (1999)
06 El Mayor-Cucapah, Mexico (2010)
07 Darfield, New Zealand (2010)

'''

print(welcome_text)

key = input("Press E key to execute program or press X to exit. ")
while True:
    if key == 'e' or key == 'E':  # '\r' and '\n' are the ASCII codes for the Enter key
        print("\nContinuing...")
        # do something to continue the program
        break
    elif key == 'x' or key == 'X': # '\x1b' is the ASCII code for the Escape key
        print("\nExiting...")
        sys.exit()
    else:
        print("\nInvalid input. Please press Enter or Escape.")
        
# Load the input text
info = []  # create an empty list to hold the lines

try:
    with open('input_info.txt', 'r') as file:
        # read the file line by line and add each line to the list
        for line in file:
            info.append(line.strip())
    print("input_info.txt found. Using current template:\n")
    for text in info:
        print(text)
            
except FileNotFoundError:
    print("input_info.txt not found. Using default template:\n")
    info = ['Project Name', 
     'Report Type', 
     'A', 
     '01 Winterfell, The North (YYYY)', 
     '02 Highgarden, The Reach (YYYY)', 
     '03 Casterly Rock, The Westerlands (YYYY)', 
     '04 The Eyrie, Vale of Arryn (YYYY)', 
     '05 Riverrun, The Riverlands (YYYY)', 
     '06 Pike, The Iron Islands (YYYY)', 
     '07 Storm\'s End, The Stormlands (2YYYY)']
    for text in info:
        print(text)

print("\nOpening Microsoft Word and Inserting Images...\n")
        
# Set up the Word application object
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = True

# Create a new document
doc = word.Documents.Add()

# Set the page size to A4
doc.PageSetup.PaperSize = win32.constants.wdPaperA4

# Remove space for all paragraphs
doc.Paragraphs.SpaceBefore = 0
doc.Paragraphs.SpaceAfter = 0

# Get the filename of your SVG file
cwd = os.getcwd()
svg_folder = cwd + '\\svg'
svg_files = sorted([f for f in os.listdir(svg_folder) if f.endswith(".svg")])

range_obj = doc.Range(0, 1)

# Loop through all svg images to insert in word
for i, file in enumerate(svg_files):
    filepath = svg_folder + '\\' + file
    
    if file.endswith('A.svg'):
        text = 'Acceleration'
    elif file.endswith('V.svg'):
        text = 'Velocity'
    elif file.endswith('D.svg'):
        text = 'Displacement'
    else:
        text = 'Error'

    # Insert the filename as italicized text
    range_obj = doc.Range(doc.Content.End - 1, doc.Content.End)
    range_obj.InsertAfter(f"{text}\n")
    #range_obj.InsertAfter(f"{file}\n")
    range_obj.Font.Italic = True  # set the Italic property of the font
    range_obj.Font.Size = 10  # set the font size to 10 points

    # Insert the SVG image at the end of the document
    range_obj = doc.Range(doc.Content.End - 1, doc.Content.End)
    shape = doc.InlineShapes.AddPicture(FileName=filepath, Range=range_obj, LinkToFile=False, SaveWithDocument=True)
    range_obj.InsertAfter(f"\n")
    shape.Width = 451
    shape.Height = 204 #

    print('Inserted ' + file)
    
    if (i + 1) == len(svg_files):
        break
    print('\nfile ends with D.svg: ' + str(file.endswith('D.svg')) + '\n')
    if file.endswith('D.svg'):
        #print('\n')
        #range_obj.InsertAfter(f"\n")
        # insert a section break after the displacement SVG image
        section = doc.Content.Sections.Add()
        #section.Range.InsertBreak(7)
        ##doc.Range(shape.Range.End, shape.Range.End).InsertBreak(win32.constants.wdSectionBreakNextPage)
    else:
        shape.Range.InsertAfter("\n")
    
def unlink_all_headers(doc):
    '''
    I need to fix this. It doesn't unlink every section all the time. 2023.06.29.
    '''
    print('\nUnlinking all sections...')
    n = 0
    section = ""
    for n, section in enumerate(doc.Sections):
        if section.Headers(win32.constants.wdHeaderFooterPrimary).LinkToPrevious:
            section.Headers(win32.constants.wdHeaderFooterPrimary).LinkToPrevious = False
#             n = n + 1
            print(repr(section) + ' Unlinked section ' + str(n))
    print('I\'m sure I unlinked it all...')
    
recheck_count = 0
while recheck_count < 10:
    unlink_all_headers(doc)
    flag = False
    for section in doc.Sections:
        if section.Headers(win32.constants.wdHeaderFooterPrimary).LinkToPrevious:
            #print('\nOops I forgot to unlink\n'+ print(repr(section) + ' Section. I will unlink again.')
            flag = True
            break
    if not flag:
        print('Now I\'m  really sure I unlinked it all...')
        break
    recheck_count += 1
    print('recheck_count_unlink:' + str(recheck_count))

print('\nGenerating section header text...')
# Collect section headers
section_header1 = []
section_header2 = []
site = ''
RP = ''
comp = ''
loc = ''
year = ''
for file in svg_files:
    if file.endswith('A.svg'):
        site = file.split()[0].split('-')[1]
        site = 'D/O' if site == 'DO' else file.split()[0].split('-')[1] #special case for 
        RP   = file.split()[1].replace('_', '-')
        comp = 'V' if file.split()[2] == 'Vert' else file.split()[2]
        loc  = info[int(RP.split('-')[1]) + 2][3:-7]
        year = info[int(RP.split('-')[1]) + 2][-5:-1]
        section_header1.append(f"[{site}{RP}] ")
        section_header2.append(f"{year} {loc}: {comp} Component")
    else:
        continue
print('\nSection header text generated.')

print('\nInserting section header text...')

# Loop through all sections in the document
for j in range(doc.Sections.Count):
    # Get the header for this section
    header = doc.Sections(j + 1).Headers(win32.constants.wdHeaderFooterPrimary).Range
    
    # Set the font size, type and style for the first set of text
    header.Text = section_header1[j]
    header.Font.Size = 16
    header.Font.Name = 'Calibri'
    header.Font.Bold = True
    
    # Insert the first set of text
    #header.InsertAfter(section_header1[i])
    
    # Move the insertion point to the end of the first set of text
    header.Collapse(win32.constants.wdCollapseEnd)
    
    # Set the font size, type and style for the second set of text
    header.Text = section_header2[j]
    header.Font.Size = 16
    header.Font.Name = 'Calibri'
    header.Font.Underline = True
      
    print(f"Inserted header {j + 1} / {len(section_header1)} : {section_header1[j]} {section_header2[j]}")

print('\nSection headers inserted.')

print('\nInserting Footer...')
    
# Add a top border to the first section's footer
footer = doc.Sections(1).Footers(win32.constants.wdHeaderFooterPrimary)
paragraphs = footer.Range.Paragraphs
first_paragraph = paragraphs(1)
border = first_paragraph.Borders(win32.constants.wdBorderTop)
border.LineStyle = win32.constants.wdLineStyleSingle

footer = doc.Sections(1).Footers(win32.constants.wdHeaderFooterPrimary).Range

# set the alignment of the footer to right-aligned
##footer.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphRight
footer.Font.Size = 8
footer.Font.Name = 'Calibri'

# Add a right aligned tab stop to the paragraph
tab_stops = footer.ParagraphFormat.TabStops
right_align_tab = tab_stops.Add(Position=1000, Alignment=win32.constants.wdAlignTabRight, Leader=win32.constants.wdTabLeaderSpaces)


# Add the footer
footer.InsertAfter(Text = f'{info[0]} \t\t P a g e | {info[2]}')
footer.Collapse(0)
footer.Fields.Add(footer,win32.constants.wdFieldPage)
footer = doc.Sections(1).Footers(win32.constants.wdHeaderFooterPrimary).Range
footer.Font.Size = 8
footer.Font.Name = 'Calibri'
footer.Collapse(0)
footer.InsertAfter(Text = f' of {info[2]}')
footer.Collapse(0)
footer.Fields.Add(footer,win32.constants.wdFieldNumPages)  
footer = doc.Sections(1).Footers(win32.constants.wdHeaderFooterPrimary).Range
footer.Font.Size = 8
footer.Font.Name = 'Calibri'
footer.Collapse(0)
footer.InsertParagraphAfter()
footer.InsertAfter(Text = f'{info[1]} \t\t')
footer.Collapse(0)
footer.Fields.Add(footer, 31, 'DATE \@ "MMMM yyyy"')
footer.Font.Size = 8
footer.Font.Name = 'Calibri'

print('\nFooter  inserted.')
# save and close the document
print('Saving document...')
# get the current date and format it as YYYY-MM-DD
today = datetime.date.today().strftime("%Y.%m.%d")
if site == 'D/O':
    site = 'DO' 
doc.SaveAs(cwd + f"\\(SHA) Annex {info[2]} - {site} Matched {RP.split('-')[0]}-Year Surface Ground Motion Time-Histories {today}.docx")
doc.Close()
print('Save success.')

# Quit the Word application
word.Quit()

input("Press any key to exit...")
sys.exit()
