import xlrd
import itertools
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
stop_words = set(stopwords.words('english'))
stop_words.add('.')
stop_words.add(',')



# reads in the pros (column 3) and cons (column 4) from the spreadsheet
def read(file_name):

    data = xlrd.open_workbook(file_name)
    sheet_no = data.sheet_by_index(0)

    pros = []
    cons = []

    for i in range(1, sheet_no.nrows):
        pros.append(sheet_no.cell(i, 3).value)
        cons.append(sheet_no.cell(i, 4).value)
            
    return pros, cons



# puts each numbered point into a separate array cell (and removes stopwords)
def separate(cell):

    # finds the indexes that have the numbers for the different points
    indexes = []

    for i in range(1, 20): # hard-coded for at most 20 points in a pro/con
        substring = "{0}.".format(i)
            
        ind = cell.find(substring)    # finding the index of occurence
        if ind == -1:
            break # breaks out of the loop as soon as a number isn't found
        else:
            indexes.append(ind)

    points = []
   
    for i in range(len(indexes)):

        start = indexes[i] + 3 # added 3 to skip over the number and the following .?

        # if this is the last point, then it goes all the way to the end of the cell
        # otherwise, it goes up to the next point
        if i+1 == len(indexes):
            end = len(cell)
        else:
           end = indexes[i+1]

        thisPoint = cell[start:end]
        points.append(removeStopwords(thisPoint))

    return points
   


def removeStopwords(sent):
    sent = sent.lower()
    word_tokens = word_tokenize(sent)
    filtered_sentence = [w for w in word_tokens if not w in stop_words]
    return filtered_sentence



def similarity(set1, set2):

    if len(set1) > len(set2):
        bigger = set1 
	smaller = set2 
    else:
        bigger = set2
	smaller = set1

    n = len(smaller.intersection(bigger))
    return n / float(len(smaller))



def compare(first_cells, sec_cells):

    matches = 0
    items = 0

    # for each cell...
    for i in range(len(first_cells)):

        # separate the contents of each cell and remove stopwords
	item1_points = separate(first_cells[i])
	item2_points = separate(sec_cells[i])

        items += len(item1_points) + len(item2_points)

        # pair-wise comparison between the items from the two cells (as a set)
        for i1 in item1_points:
            for i2 in item2_points:

                # convert the items to a set
	        set1 = set(i1)
	        set2 = set(i2)

                # compare the items using Jaccard similarity
                sim = similarity(set1, set2)

                # if the similarity is greater than the threshold
                # stop comparing and increment the match count
                if sim > 0.25:
                    matches += 1
                    break

    # return the number of matches and the number of items 
    return matches, items



def compareParticipants(filename1, filename2):

    [pros1, cons1] = read(filename1)
    [pros2, cons2] = read(filename2)
   
    [proMatches, proItems] = samePros = compare(pros1, pros2)
    [conMatches, conItems] = sameCons = compare(cons1, cons2)

    # Jaccard similarity
    sim = (proMatches + conMatches) / float(proItems + conItems - proMatches - conMatches) 

    return sim



files = [
    "ThesisEventDetailedSummary_P1.xlsx", 
    "ThesisEventDetailedSummary_P2.xlsx", 
    "ThesisEventDetailedSummary_P3.xlsx", 
    "ThesisEventDetailedSummary_P4.xlsx"
]

#files = [
#    "test1.xlsx", 
#    "test2.xlsx", 
#]

agreements = []

#sim = round(compareParticipants(files[0], files[0]), 2)
#print "sim = ", sim

for i in range(len(files)):
    for j in range(len(files)):
        sim = round(compareParticipants(files[i], files[j]), 2)
        agreements.append(sim)

print "Agreement Score Matrix is: "
for i in range(len(files) * len(files)):
    if i % len(files) < len(files)-1:
	print agreements[i], "\t",
    else:
	print agreements[i]
