import os
import unittest
import zipfile
from StringIO import StringIO

import ooolib
from lxml import etree
from ooolib.tests.utils import prepare_mkdtemp


class TestCell(unittest.TestCase):

    def setUp(self):
        self.dirname = prepare_mkdtemp(self)

    def test_cell_border(self):
        # create document
        thin_value = '0.001in solid #ff0000'
        bold_value = '0.01in solid #00ff00'

        filename = os.path.join(self.dirname, "test.odt")

        doc = ooolib.Calc("test_cell")

        doc.set_cell_property('border', thin_value)
        doc.set_cell_value(2, 2, "string", "Name")

        doc.set_cell_property('border', bold_value)
        doc.set_cell_value(2, 4, "string", "Value")

        doc.save(filename)

        # check created document
        handle = zipfile.ZipFile(filename)
        xdoc = etree.parse(StringIO(handle.read('content.xml')))

        namespaces = {
            'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
            'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
            'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
            'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
        }
        border_thin = xdoc.xpath('//style:table-cell-properties[@fo:border="%s"]/../@style:name' % thin_value,
                                 namespaces=namespaces)
        self.assertEqual(len(border_thin), 1)

        border_bold = xdoc.xpath('//style:table-cell-properties[@fo:border="%s"]/../@style:name' % bold_value,
                                 namespaces=namespaces)
        self.assertEqual(len(border_bold), 1)
        cell_style = xdoc.xpath('//text:p[text()="Name"]/../@table:style-name', namespaces=namespaces)
        self.assertEqual(cell_style, border_thin)

        cell_style = xdoc.xpath('//text:p[text()="Value"]/../@table:style-name', namespaces=namespaces)
        self.assertEqual(cell_style, border_bold)
