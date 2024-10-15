import xml.etree.ElementTree as ET


ns = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
}

tree = ET.parse('decka-1/content.xml')
root = tree.getroot()

# import pdb; pdb.set_trace()  # !!!
for sheet in root.findall('office:body/office:spreadsheet', ns):
    # print("sheet:", sheet)
    for row_position, row in enumerate(sheet.findall('table:table/table:table-row', ns)):
        # print("=" * 60)
        # print("row:", row)
        for cell_position, cell in enumerate(row.findall('table:table-cell[@office:value-type]', ns)):
            # print("-" * 60)
            # print("cell:", cell, cell.attrib)
            paragraph = cell.find("text:p", ns)
            if paragraph is not None:
                print(row_position, cell_position, paragraph.text)

body = ET.tostring(root, encoding='utf-8', xml_declaration=True)

print(body.decode("utf-8"))

# xml.etree.ElementTree.tostring(element, encoding='us-ascii', method='xml', *, xml_declaration=None, default_namespace=None, short_empty_elements=True)
# xml.etree.ElementTree.tostringlist(element, encoding='us-ascii', method='xml', *, xml_declaration=None, default_namespace=None, short_empty_elements=True)
# print(etree.tostring(my_doc, pretty_print=True, xml_declaration=True, encoding="utf-8"))
# print(etree.tostring(command, pretty_print=True, xml_declaration=True, encoding="utf-8", standalone=False))

"""
  <office:body>
    <office:spreadsheet>
      <table:calculation-settings table:automatic-find-labels="false" table:use-regular-expressions="false" table:use-wildcards="true"/>
      <table:table table:name="List1" table:style-name="ta1">
        <table:table-column table:style-name="co1" table:number-columns-repeated="3" table:default-cell-style-name="Default"/>
        <table:table-row table:style-name="ro1">
          <table:table-cell office:value-type="string" calcext:value-type="string">
            <text:p>Matěj</text:p>
          </table:table-cell>
          <table:table-cell table:number-columns-repeated="2"/>
        </table:table-row>
"""
