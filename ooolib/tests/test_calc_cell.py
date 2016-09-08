import os
import unittest
import zipfile
from StringIO import StringIO

import ooolib
from lxml import etree
from ooolib.tests.utils import prepare_mkdtemp


class TestCell(unittest.TestCase):

    namespaces = {
        'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
        'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
        'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
        'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
    }

    def setUp(self):
        self.dirname = prepare_mkdtemp(self)

    def test_cell_border(self):
        # create odt document
        thin_value = '0.001in solid #ff0000'
        bold_value = '0.01in solid #00ff00'

        filename = os.path.join(self.dirname, "test.ods")
        doc = ooolib.Calc("test_cell")
        doc.set_cell_property('border', thin_value)
        doc.set_cell_value(2, 2, "string", "Name")
        doc.set_cell_property('border', bold_value)
        doc.set_cell_value(2, 4, "string", "Value")
        doc.save(filename)

        # check created document
        handle = zipfile.ZipFile(filename)
        xdoc = etree.parse(StringIO(handle.read('content.xml')))

        style_definition = xdoc.xpath('//style:table-cell-properties[@fo:border="%s"]/../@style:name' % thin_value,
                                      namespaces=self.namespaces)
        cell_style = xdoc.xpath('//text:p[text()="Name"]/../@table:style-name', namespaces=self.namespaces)
        self.assertEqual(cell_style, style_definition)

        style_definition = xdoc.xpath('//style:table-cell-properties[@fo:border="%s"]/../@style:name' % bold_value,
                                      namespaces=self.namespaces)
        cell_style = xdoc.xpath('//text:p[text()="Value"]/../@table:style-name', namespaces=self.namespaces)
        self.assertEqual(cell_style, style_definition)

    def test_cell_wrap_option(self):
        # create odt document
        filename = os.path.join(self.dirname, "test.ods")
        doc = ooolib.Calc("test_cell")
        doc.set_cell_property('wrap-option', 'wrap')
        doc.set_cell_value(2, 2, "string", "Name")
        doc.save(filename)

        # check created document
        handle = zipfile.ZipFile(filename)
        xdoc = etree.parse(StringIO(handle.read('content.xml')))

        style_definition = xdoc.xpath('//style:table-cell-properties[@fo:wrap-option="wrap"]/../@style:name',
                                      namespaces=self.namespaces)
        cell_style = xdoc.xpath('//text:p[text()="Name"]/../@table:style-name', namespaces=self.namespaces)
        self.assertEqual(cell_style, style_definition)

    def test_cell_hyphenate(self):
        # create odt document
        filename = os.path.join(self.dirname, "test.ods")
        doc = ooolib.Calc("test_cell")
        doc.set_cell_property('hyphenate', True)
        doc.set_cell_value(2, 2, "string", "Name")
        doc.save(filename)

        # check created document
        handle = zipfile.ZipFile(filename)
        xdoc = etree.parse(StringIO(handle.read('content.xml')))

        style_definition = xdoc.xpath('//style:text-properties[@fo:hyphenate="true"]/../@style:name',
                                      namespaces=self.namespaces)
        cell_style = xdoc.xpath('//text:p[text()="Name"]/../@table:style-name', namespaces=self.namespaces)
        self.assertEqual(cell_style, style_definition)
