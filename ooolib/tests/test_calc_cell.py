# -*- coding: utf-8 -*-
from __future__ import unicode_literals

import os
import unittest
import zipfile
import six

from lxml import etree

import ooolib
from ooolib.tests.utils import prepare_mkdtemp

if six.PY3:
    from io import BytesIO as BufferIO
else:
    from StringIO import StringIO as BufferIO


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
        xdoc = etree.parse(BufferIO(handle.read('content.xml')))

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
        xdoc = etree.parse(BufferIO(handle.read('content.xml')))

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
        xdoc = etree.parse(BufferIO(handle.read('content.xml')))

        style_definition = xdoc.xpath('//style:text-properties[@fo:hyphenate="true"]/../@style:name',
                                      namespaces=self.namespaces)
        cell_style = xdoc.xpath('//text:p[text()="Name"]/../@table:style-name', namespaces=self.namespaces)
        self.assertEqual(cell_style, style_definition)

    def test_cell_padding(self):
        # create odt document
        filename = os.path.join(self.dirname, "test.ods")
        doc = ooolib.Calc("test_cell")
        doc.set_cell_property('padding-left', '0.1in')
        doc.set_cell_value(2, 2, "string", "Left")
        doc.set_cell_property('padding-left', False)

        doc.set_cell_property('padding-right', '0.2in')
        doc.set_cell_value(3, 2, "string", "Right")
        doc.set_cell_property('padding-right', False)

        doc.set_cell_property('padding', '0.3in')
        doc.set_cell_value(4, 2, "string", "Full")
        doc.set_cell_property('padding', False)

        doc.set_cell_value(5, 2, "string", "No-padding")

        doc.save(filename)

        # check created document
        handle = zipfile.ZipFile(filename)
        xdoc = etree.parse(BufferIO(handle.read('content.xml')))

        style_definition = xdoc.xpath('//style:table-cell-properties[@fo:padding-left="0.1in"]/../@style:name',
                                      namespaces=self.namespaces)
        cell_style = xdoc.xpath('//text:p[text()="Left"]/../@table:style-name', namespaces=self.namespaces)
        self.assertEqual(cell_style, style_definition)

        style_definition = xdoc.xpath('//style:table-cell-properties[@fo:padding-right="0.2in"]/../@style:name',
                                      namespaces=self.namespaces)
        cell_style = xdoc.xpath('//text:p[text()="Right"]/../@table:style-name', namespaces=self.namespaces)
        self.assertEqual(cell_style, style_definition)

        style_definition = xdoc.xpath('//style:table-cell-properties[@fo:padding="0.3in"]/../@style:name',
                                      namespaces=self.namespaces)
        cell_style = xdoc.xpath('//text:p[text()="Full"]/../@table:style-name', namespaces=self.namespaces)
        self.assertEqual(cell_style, style_definition)

        cell_style = xdoc.xpath('//text:p[text()="No-padding"]/../@table:style-name', namespaces=self.namespaces)
        self.assertEqual(cell_style, [])

    def test_set_cell_formula_value(self):
        doc = ooolib.Calc("test_cell")
        doc.set_cell_value(1, 1, "float", 1)
        doc.set_cell_value(2, 1, "float", 2)
        doc.set_cell_value(3, 1, "formula", "=A1+B1", 3)
        self.assertEqual(doc.sheets[0].sheet_values[(3, 1)]['formula_value'], 3)

    def test_default_cell_formula_value(self):
        doc = ooolib.Calc("test_cell")
        doc.set_cell_value(1, 1, "float", 1)
        doc.set_cell_value(2, 1, "float", 2)
        doc.set_cell_value(3, 1, "formula", "=A1+B1")
        self.assertEqual(doc.sheets[0].sheet_values[(3, 1)]['formula_value'], "0")


class TestCalcSheet(unittest.TestCase):

    def test_clean_formula(self):
        sheet = ooolib.CalcSheet('Sheet Name')

        data = sheet.clean_formula('The test.')
        self.assertEqual(data, 'The test.')

        data = sheet.clean_formula('=SUM(A1:A2)')
        self.assertEqual(data, 'oooc:=SUM([.A1]:[.A2])')

        data = sheet.clean_formula('=IF((A5>A4);A3;"The test.")')
        self.assertEqual(data, 'oooc:=IF(([.A5]&gt;[.A4]);[.A3];&quot;The test.&quot;)')


class Bird(object):

    def __init__(self, name):
        self.name = name

    def __str__(self):
        return "bird {}".format(self.name)


class TestUnicode(unittest.TestCase):

    def test_sheet_name(self):
        doc = ooolib.Calc('Žluťoučký kůň')
        self.assertEqual(doc.sheets[0].sheet_name, 'Žluťoučký kůň')

    def test_sheet_datatype_and_value(self):
        doc = ooolib.Calc('Žluťoučký kůň')
        doc.set_cell_value(2, 2, 'šílený', 'čížek')
        self.assertEqual(doc.get_cell_value(2, 2), ('string', 'čížek'))

    def test_sheet_datatype_and_value_object(self):
        doc = ooolib.Calc('Žluťoučký kůň')
        doc.set_cell_value(2, 2, 'string', Bird('čížek'))
        self.assertEqual(doc.get_cell_value(2, 2), ('string', 'bird čížek'))

    def test_clean_formula_apos(self):
        doc = ooolib.Calc('Žluťoučký kůň')
        sheet = doc.sheets[0]
        data = sheet.clean_formula("=UNICODE('Žluťoučký kůň příšerně úpěl ďábelské ódy.')")
        self.assertEqual(data, "oooc:=UNICODE(&apos;Žluťoučký kůň příšerně úpěl ďábelské ódy.&apos;)")

    def test_clean_formula_quot(self):
        doc = ooolib.Calc('Žluťoučký kůň')
        sheet = doc.sheets[0]
        data = sheet.clean_formula('=UNICODE("Žluťoučký kůň příšerně úpěl ďábelské ódy.")')
        self.assertEqual(data, "oooc:=UNICODE(&quot;Žluťoučký kůň příšerně úpěl ďábelské ódy.&quot;)")


class TestGetCellContent(unittest.TestCase):

    def test_string(self):
        doc = ooolib.Calc('Test')
        doc.set_cell_value(1, 1, 'string', 'text')
        self.assertEqual(doc.get_cell_value(1, 1), ('string', 'text'))

    def test_link(self):
        doc = ooolib.Calc('Test')
        doc.set_cell_value(1, 1, 'link', ('url', 'label'))
        self.assertEqual(doc.get_cell_links(1, 1), [('url', 'label')])

    def test_annotation(self):
        doc = ooolib.Calc('Test')
        doc.set_cell_value(1, 1, 'annotation', 'foo')
        self.assertEqual(doc.get_cell_annotation(1, 1), ('annotation', 'foo'))


class TestParseContent(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        path = os.path.join(os.path.dirname(__file__), "fixtures/test-cells.ods")
        cls.doc = ooolib.Calc(opendoc=path)

    def test_content_parse(self):
        self.assertEqual(self.doc.get_cell_value(1, 1), ('string', 'text'))
        self.assertIsNone(self.doc.get_cell_links(1, 1))
        self.assertIsNone(self.doc.get_cell_annotation(1, 1))

        self.assertEqual(self.doc.get_cell_value(1, 2), ('float', '42'))
        self.assertEqual(self.doc.get_cell_value(1, 3), ('string', 'link: https://www.nic.cz/'))
        self.assertEqual(self.doc.get_cell_value(1, 4), ('string', 'link with label: https://www.mojeid.cz/'))

        self.assertEqual(self.doc.get_cell_links(1, 3), [('https://www.nic.cz/', 'https://www.nic.cz')])
        self.assertEqual(self.doc.get_cell_links(1, 4), [('https://www.mojeid.cz/', 'MojeID')])

        self.assertEqual(self.doc.get_cell_value(1, 5), ('string', 'Comment.'))
        self.assertEqual(self.doc.get_cell_value(1, 6), (
            'string', 'Cell with two links: https://www.nic.cz/ and https://www.mojeid.cz/.'))
        self.assertEqual(self.doc.get_cell_links(1, 6), [
            ('https://www.nic.cz/', 'CZ.NIC'),
            ('https://www.mojeid.cz/', 'MojeID'),
        ])
