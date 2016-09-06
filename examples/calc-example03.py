#!/usr/bin/python

import sys
sys.path.append('..')
import ooolib

# Create the document
doc = ooolib.Calc()

# Document Properties
doc.set_meta('title', 'The Search')
doc.set_meta('subject', 'Searching for the Grail')
doc.set_meta('description', 'This document is all about finding the grail.')

# Set meta data for the document
doc.set_meta('creator', 'King Arther')
doc.set_meta('editor', 'Sir Robin')

# User Defined Meta Data Names
doc.set_meta('user1name', 'Name')
doc.set_meta('user2name', 'Capital of Assyria')
doc.set_meta('user3name', 'Favourite colour')
doc.set_meta('user4name', 'Air-speed velocity of an unladen swallow')

# User Defined Meta Data Values
doc.set_meta('user1value', 'Sir Lancelot of Camelot')
doc.set_meta('user2value', "I don't know that")
doc.set_meta('user3value', 'Blue')
doc.set_meta('user4value', 'What do you mean? An African or European swallow?')

# Single Cell
doc.set_cell_value(1, 1, "string", "To see the changes select: File -> Properties...")

# Write out the document
doc.save("calc-example03.ods")
