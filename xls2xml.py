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

def _has_group(book):
    return True if book.nsheets > 1 else False


def _get_col_index(sheet_index, headerdict, colname):
    return headerdict[sheet_index].index(colname)


def _id_in_group(book, sheet_index, _index):
    pass


def gen_xml(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)
    sheetnames = book.sheet_names()

    # Get the first worksheet and column names
    data_sheet0 = book.sheet_by_index(0)
    colnames = data_sheet0.row_values(0)

    # Create dict of headers
    headerdict = {}
    for i in range(0, book.nsheets):
        repeat_sheet = book.sheet_by_index(i)
        data = [repeat_sheet.cell_value(r,c) for c in range(repeat_sheet.ncols) for r in range(1)]
        headerdict[i] = data
    print 'headerdict', headerdict

    # Identify key column header indices
    version_col_index =  _get_col_index(0, headerdict, '__version__')
    _uuid_col_index = _get_col_index(0, headerdict, '_uuid')

    # generate variables for sample group sheet
    if _has_group(book):
        _index_col_index =             _get_col_index(0, headerdict, '_index')
        _parent_table_name_col_index = _get_col_index(1, headerdict, '_parent_table_name')
        _parent_index_col_index =      _get_col_index(1, headerdict, '_parent_index')

        group_sheet1 = book.sheet_by_index(1)

        # Create data structure to store lists of row indices keyed by _parent_index
        # This will likely eventually go in a loop
        group1_indices = {}
        for i, v in enumerate(group_sheet1.col_values(_parent_index_col_index)):
            if v in group1_indices:
                group1_indices[v].append(i)
            else:
                group1_indices[v] = [i]

        print group1_indices


    # Loop through rows in first datasheet to output XML
    for row in range(1, data_sheet0.nrows):
        # column names, __version__ column index, and version value
        version = data_sheet0.cell_value(1,version_col_index)

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
        for i, colname in enumerate(headerdict[0][:version_col_index]):
            colname_el = ET.SubElement(root, colname)
            colname_el.text = str(data_sheet0.cell_value(row,i))

        # Begin work on repeating fields.
        if _has_group(book):
            # note slice, to avoid including the datasheet; To reference sheet number, use enum + 1
            for enum, gsheet in enumerate(sheetnames[1:]):
                print enum+1, gsheet, "............."
                repeat_sheet = book.sheet_by_name(gsheet) # in this example: group_cooking
                print "from data_sheet0", data_sheet0.cell_value(enum, _index_col_index), "+++"

                #this assumes all the elements of the row will include the same parent table name
                if data_sheet0.name == repeat_sheet.cell_value(1,_parent_table_name_col_index):
                    print "===", repeat_sheet.cell_value(1,_parent_table_name_col_index), "==="
                    print "paired sheet", enum, "--->", enum+1
                    print group1_indices
                    print repeat_sheet.cell_value(1,_parent_index_col_index), "repeat_sheet.cell_value(1,_parent_index_col_index)"

                    # print data row index value
                    index_key = data_sheet0.cell_value(row, _index_col_index)
                    print index_key, "---index key"

                    if index_key in group1_indices:
                        print group1_indices[index_key]
                        for group_row in range(len(group1_indices[index_key])):
                            sheet_el = ET.SubElement(root, gsheet)
                            # TODO
                            # sheet_el.value = repeat_sheet._cell_value(group_row,

                    print round(index_key,1)
                    print '__________'


        # create __version__ element
        version_el = ET.SubElement(root,"__version__")
        version_el.text = str(data_sheet0.cell_value(row,version_col_index))

        # create meta element with nested instanceID
        meta_el = ET.SubElement(root,"meta")
        instance_ID_el = ET.SubElement(meta_el, "instanceID")
        iID = data_sheet0.cell_value(row, _uuid_col_index)
        instance_ID_el.text = iID if len(iID) > 0 else str(uuid.uuid4())

        #create the xml files
        tree = ET.ElementTree(root)
        output_fn = instance_ID_el.text + '.xml'
        tree.write(output_fn, pretty_print=True, xml_declaration=True,   encoding="utf-8")

if __name__ == "__main__":
    INPUT_EXCEL_FILE = sys.argv[1] # "xls-to-xml-test.xlsx"
    KPI_UID = sys.argv[2]
    KC_UUID = sys.argv[3]

    gen_xml(INPUT_EXCEL_FILE)
