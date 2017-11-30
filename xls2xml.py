#!/usr/bin/env python

import xlrd
import sys
import uuid
from lxml import etree as ET

"""
-- Sample command lines --
python xls2xml.py xls-to-xml-test.xlsx aZCyzqYa2aqEtf2945cna6 524fc08b8a0e4d8d857dded88d5fb882
python xls2xml.py xls-to-xml-repeats.xlsx aZCyzqYa2aqEtf2945cna6 524fc08b8a0e4d8d857dded88d5fb882

-- Sample XML output (no repeats) --
<?xml version="1.0" ?>
<aZCyzqYa2aqEtf2945cna6 xmlns:jr="http://openrosa.org/javarosa" xmlns:orx="http://openrosa.org/xforms" id="aZCyzqYa2aqEtf2945cna6" version="vU6wm35wT9e5Bpr2kf5Jyw">
          <formhub>
            <uuid>524fc08b8a0e4d8d857dded88d5fb882</uuid>
          </formhub>
          <age>100</age>
          <happiness>yes</happiness>
          <name>Theodore</name>
          <date>2017-11-27</date>
          <__version__>vU6wm35wT9e5Bpr2kf5Jyw</__version__>
          <meta>
            <instanceID>ea0a35b1-125d-4856-ab09-087e38c6131b</instanceID>
          </meta>
        </aZCyzqYa2aqEtf2945cna6>

-- Sample XML output (with repeats) --
<?xml version="1.0" ?>
<ayJS36BXDhJRsCsXWmRiPu xmlns:jr="http://openrosa.org/javarosa" xmlns:orx="http://openrosa.org/xforms" id="ayJS36BXDhJRsCsXWmRiPu" version="vZSdJcLdcv3vuwdgiNeSeR">
          <formhub>
            <uuid>ca8356db0ef6463685af4b071c537722</uuid>
          </formhub>
          <start>2017-11-29T16:10:15.000-05:00</start>
          <end>2017-11-29T16:12:55.000-05:00</end>
          <Name>Minnie Mouse</Name>
          <Birthdate>1928-10-18</Birthdate>
          <age>89</age>
          <happiness>yes</happiness>
          <group_cooking>
            <Cooking_Equipment>Pan</Cooking_Equipment>
            <Years_Owned>78</Years_Owned>
          </group_cooking><group_cooking>
            <Cooking_Equipment>Ladle</Cooking_Equipment>
            <Years_Owned>77</Years_Owned>
          </group_cooking>
          <__version__>vZSdJcLdcv3vuwdgiNeSeR</__version__>
          <meta>
            <instanceID>uuid:705bd87e-c2ee-4149-b6fa-b459e5213852</instanceID>
          </meta>
        </ayJS36BXDhJRsCsXWmRiPu>
"""

NSMAP = {"jr" :  'http://openrosa.org/javarosa',
         "orx" : 'http://openrosa.org/xforms'}



def gen_xml(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)
    sheetnames = book.sheet_names()

    # get the first worksheet
    data_sheet1 = book.sheet_by_index(0)

    for row in range(1, data_sheet1.nrows):
        # column names, __version__ column index, and version value
        colnames = data_sheet1.row_values(0)
        version_col_index = colnames.index("__version__")
        version = data_sheet1.cell(1,version_col_index).value

        # create root element
        root = ET.Element(KPI_UID, nsmap = NSMAP)
        root.set('id', KPI_UID)
        root.set("version", version)
        # create formhub element with nested uuid
        fhub_el = ET.SubElement(root, "formhub")
        kc_uuid_el = ET.SubElement(fhub_el, "uuid")
        kc_uuid_el.text = KC_UUID

        # create elements from the first column up to and including the _version__
        slice_index = version_col_index + 1
        for i, colname in enumerate(colnames[:version_col_index]):
            colname_el = ET.SubElement(root, colname)
            colname_el.text = str(data_sheet1.cell(row,i).value)

        # begin work on repeating fields
        if _has_group(book):
            for sheet in sheetnames[1:]:
                pass
                # if _id_in_group(book, sheet_index, _index):
                # sheet_el = ET.SubElement(root, sheet)

        # create __version__ element
        version_el = ET.SubElement(root,"__version__")
        version_el.text = str(data_sheet1.cell(row,version_col_index).value)

        # create meta element with nested instanceID
        meta_el = ET.SubElement(root,"meta")
        instance_ID_el = ET.SubElement(meta_el, "instanceID")
        _uuid_col_index = colnames.index("_uuid")
        iID = data_sheet1.cell(row, _uuid_col_index).value
        instance_ID_el.text = iID if len(iID) > 0 else str(uuid.uuid4())

        #create the xml files
        tree = ET.ElementTree(root)
        output_fn = instance_ID_el.text + '.xml'
        tree.write(output_fn, pretty_print=True, xml_declaration=True,   encoding="utf-8")

def _has_group(book):
    return True if book.nsheets > 1 else False

def _id_in_group(book, sheet_index, _index):
    pass


if __name__ == "__main__":
    INPUT_EXCEL_FILE = sys.argv[1] # "xls-to-xml-test.xlsx"
    KPI_UID = sys.argv[2]
    KC_UUID = sys.argv[3]

    gen_xml(INPUT_EXCEL_FILE)
