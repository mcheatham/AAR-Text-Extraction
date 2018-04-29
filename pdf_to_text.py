import re
import os
import pandas
import PyPDF2

base_dir = "AAR documents"

# paragraph indicators
# a tab
# period, some spaces, a newline, possibly some spaces, capital letter
# period, at least three spaces, capital letter
paragraph1 = re.compile(r'\t')
paragraph2 = re.compile(r'\.[ ]+[\n\r]+[ ]*([A-Z])')
paragraph3 = re.compile(r'\s\s\s[\s]*([A-Z])')

# kludge to correct cases where we have a capitalized word, a newline, 
# and then another capitalized word -> remove the newline
mistakes = re.compile(r'([A-Z][a-z]*)AAR_PARA([A-Z][a-z]*)')

# for removing all newlines
newline = re.compile(r'[\n\r]+')

# for removing image encodings
special_chars = re.compile(r'[#!+()%$&]')
all_chars = re.compile(r'[\w]')


for pdf in os.listdir(base_dir):

    if pdf.endswith(".pdf"):

        print (pdf)

        txt = pdf.replace('.pdf', '.txt')

        output = open(os.path.join(base_dir, txt), 'w')
        
        contents = ''

        pdf_file = open(os.path.join(base_dir, pdf), 'rb')
        read_pdf = PyPDF2.PdfFileReader(pdf_file)

        for j in range(0, read_pdf.getNumPages()):

            page = read_pdf.getPage(j).extractText()

            # first try figuring out the paragraphs
            taggedParagraphs = paragraph1.sub('AAR_PARA', page)
            taggedParagraphs = paragraph2.sub(r'. AAR_PARA\1', taggedParagraphs)
            taggedParagraphs = paragraph3.sub(r'AAR_PARA\1', taggedParagraphs)

            # correct mistakes where paragraph breaks were added in the 
            # middle of proper name phrases
            corrected = mistakes.sub(r'\1 \2', taggedParagraphs)

            # now remove all newlines
            noNewlines = newline.sub('', corrected)

            # restore the paragraphs
            paragraphs = noNewlines.replace('AAR_PARA', '\n\n')

            # don't break on semi-colons
            paragraphs = paragraphs.replace(';\n\n', '; ')

            # only add a paragraph if it's not mostly special characters
            numSpecial = len(special_chars.findall(paragraphs))
            numChars = len(all_chars.findall(paragraphs))

            if numChars > 0 and numSpecial / numChars < 0.3:
                contents += paragraphs

            # TODO try to handle paragraph breaks across pages, keeping 
            # in mind page headers and page numbers
        
        output.write(contents)
        output.close()

