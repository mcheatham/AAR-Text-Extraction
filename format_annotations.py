import os
import xlrd
import pandas
import re

annotationsDir = "AAR annotations"
documentsDir = "AAR documents"

one = False

for f in os.listdir(annotationsDir):

    if f.endswith(".xlsx") and f.startswith("P"):

        one = True

        path = os.path.join(annotationsDir, f)
        print(path)
        annotationFile = pandas.read_excel(path)

        # print(annotationFile.columns)

        strengths = annotationFile['Pros/Successful tasks'].values
        print(strengths)

        weaknesses = annotationFile['Cons/Recommendations'].values
        # print(weaknesses)

        # we to output a file with the document id number
        # and the strengths and weaknesses for that document

        # TODO figure out which document each row is for
        for i in range(0, len(strengths)):

            # look through each of the .txt files in the docs dir
            found = False

            for doc in os.listdir(documentsDir):

                if doc.endswith(".txt"):

                    docPath = os.path.join(documentsDir, doc)
                    with open(docPath, 'r') as docFile:

                        docContents = docFile.read()

                        if strengths[i] in docContents or weaknesses[i] in docContents:

                            #print(i, " is ", doc)
                        #    print(j + " " + strengths[i] + " " + weaknesses[i])
                            found = True
                            break

            #if not found:
                #print(i, " not found in the documents")
    if one:
        break
