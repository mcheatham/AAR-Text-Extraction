import os
import xlrd
import re

annotationDir = "AAR annotations"
reportDir = "AAR documents"

#for f in os.listdir(annotationDir):
#
#    if f.endswith(".xlsx") and f.startswith("P"):
#
#        path = os.path.join(annotationDir, f)
#        # print(path)
#        annotationFile = pandas.read_excel(path)
#
#        # print(annotationFile.columns)
#
#        strengths = annotationFile['Pros/Successful tasks'].values
#        # print(strengths)
#
#        weaknesses = annotationFile['Cons/Recommendations'].values
#        # print(weaknesses)
#
#        # we to output a file with the document id number, 
#        # and the strengths and weaknesses for that document, 
#        # separated into individual sentences
#
#        # TODO figure out which document each row is for
#        for i in range(0, len(annotationFile.count())):
#
                        #if strengths[i] in page_content or weaknesses[i] in page_content:
                        #    print(j + " " + strengths[i] + " " + weaknesses[i])
                        #    break
