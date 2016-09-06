#!/usr/bin/python

import sys
sys.path.append('..')
import ooolib

# See if there is a document to open on the command line
if len(sys.argv) != 2:
    print "Usage:\n\t%s FILENAME.ods\n" % sys.argv[0]
    sys.exit(0)

# Use opendoc to select the document to edit.
doc = ooolib.Calc(opendoc=sys.argv[1])

# Now that the document has been opened and the data loaded,
# we can add to the document, or display information from the
# document.

# First we display current contents
print "Current Data"
print " Meta Values"
metalist = ['creator', 'editor', 'title', 'subject', 'description',
	    'user1name', 'user2name', 'user3name', 'user4name',
	    'user1value', 'user2value', 'user3value', 'user4value',
	    'keyword']
for metaname in metalist:
	value = doc.get_meta_value(metaname)
	print "  %s=%s" % (metaname, value)

print " Content Values"
# We need to walk through all of the sheets
count = doc.get_sheet_count()
for c in range(0, count):
    doc.set_sheet_index(c)
    name = doc.get_sheet_name()
    print "  Sheet Name: %s" % name
    (cols, rows) = doc.get_sheet_dimensions()
    for row in range(1, rows + 1):
        print "  ",
        for col in range(1, cols + 1):
            data = doc.get_cell_value(col, row)
            print data,
        print

# Second we add to the document
print "New Data"
print " Meta Data"
print "  Reset title"
doc.set_meta("title", "She turned me into a newt")

print " Content Values"
print "  Setting A1 to '3' on Sheet1"
doc.set_sheet_index(0)
doc.set_cell_value(1, 1, 'float', 3)

# Finally, we write out the modified document
print "Saving Document as calc-example09.ods"
doc.save("calc-example09.ods")
