#!/usr/bin/env python

import xlrd
import sys
import uuid
from lxml import etree as ET

"""
-- Sample commands --
xls2xml.py xls-to-xml-test.xlsx aZCyzqYa2aqEtf2945cna6 524fc08b8a0e4d8d857dded88d5fb882
xls2xml.py BasicRepeatForm.xlsx vp8Y2sBnY5csBwReweRMGU de70c38d99ce4258bfb70fed1b8f1efa

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


def _has_group(book):
    return True if book.nsheets > 1 else False


def _get_col_index(sheet_index, headers, colname):
    return headers[sheet_index].index(colname)


def _gen_headers(book):
    """Create dict of headers, keyed by sheet index.

    Sample dict:
    {0: [u'start', u'end', u'Name', u'Birthdate', u'age', u'happiness',
         u'__version__', u'_id', u'_uuid', u'_submission_time', u'_index', u''],
     1: [u'Cooking Equipment', u'Years Owned', u'_index',
         u'_parent_table_name', u'_parent_index']}
    """
    headers = {}
    for i in range(0, book.nsheets):
        repeat_sheet = book.sheet_by_index(i)
        data = [repeat_sheet.cell_value(r,c) for c in range(repeat_sheet.ncols) for r in range(1)]
        headers[i] = data
    return headers

def _gen_group_indices(book, sheet_index):
    """
    Create data structure to store lists of row indices keyed by _parent_index.

    Sample dict:
    {
        2.0: [1],
        3.0: [2, 3],
        4.0: [4, 5],
        5.0: [6]
    }
    """
    colnames = _gen_headers(book)[sheet_index]
    g_parent_index_col_index = colnames.index("_parent_index")

    group_indices = {}

    group_sheet = book.sheet_by_index(sheet_index)
    for group_row, parent_row in enumerate(group_sheet.col_values(g_parent_index_col_index)):
        if parent_row in group_indices:
            group_indices[parent_row].append(group_row)
        else:
            if str(parent_row).encode('ascii') == '_parent_index':
                continue
            group_indices[parent_row] = [group_row]

    return group_indices

def _gen_group_index_list(book):
    """
    Create list of group indices keyed by sheetname

    Sample dict:
    [{
    u 'group_cooking': {
        2.0: [1],
        3.0: [2, 3],
        4.0: [4, 5],
        5.0: [6]
    }
}, {
    u 'pets': {
        2.0: [1],
        3.0: [2],
        4.0: [3, 4, 5]
    }
}]
    """

    group_index_list = []
    groups_indices = {}
    for s_i, sheetname in enumerate(book.sheet_names()):
        if s_i == 0:
            pass
        else:
            groups_indices[sheetname] = _gen_group_indices(book, s_i)
            group_index_list.append(groups_indices)
            groups_indices = {}
    return group_index_list


def _parse_multi_select_data(multi_selects, header, text):
    """ Look at data and if the option is selected, add to multi_selects """
    if text == '1':
        name, value = header.split('/')
        if name not in multi_selects:
            multi_selects[name] = []
        multi_selects[name].append(value)

def _gen_multi_selects(parent_node, multi_selects):
    """ Generate XML elements from multi_selects """
    for key, value in multi_selects.items():
        colname_el = ET.SubElement(parent_node, key)
        colname_el.text = " ".join(value)

def _gen_xml_elements0(book, headers, row):
    """
    Create elem dict, which will populate the beginning of the XML file
    """
    data_sheet0 = book.sheet_by_index(0)

    version_col_index =  _get_col_index(0, headers, '__version__')
    version = data_sheet0.cell_value(1,version_col_index)

    NSMAP = {"jr" :  'http://openrosa.org/javarosa',
         "orx" : 'http://openrosa.org/xforms'}

    # create root element
    root = ET.Element(KPI_UID, nsmap = NSMAP)
    root.set('id', KPI_UID)
    root.set("version", version)

    # create formhub element with nested uuid
    fhub_el = ET.SubElement(root, "formhub")
    kc_uuid_el = ET.SubElement(fhub_el, "uuid")
    kc_uuid_el.text = KC_UUID

    # create elements from the first column up to and including the _version__
    multi_selects = {}
    for i, colname in enumerate(headers[0][:version_col_index]):
        text0 = data_sheet0.cell_value(row,i)
        if "/" in colname:
            _parse_multi_select_data(multi_selects, colname, text0)
        else:
            colname_el = ET.SubElement(root, colname)
            colname_el.text = text0

    _gen_multi_selects(root, multi_selects)

    elems = {row: {}}
    elems[row]['root'] = root

    return elems

def _gen_group_detail(book, row, headers, data_sheet0, root):
        group_index_list = _gen_group_index_list(book)
        _index_col_index = _get_col_index(0, headers, '_index')

        for sheet_i, g_indices in enumerate(group_index_list):
            sheet_i += 1  # add 1 because group_index_list doesn't include the first worksheet
            group_sheet = book.sheet_by_index(sheet_i)
            group_sheetname = group_sheet.name
            _parent_index = data_sheet0.cell_value(row, _index_col_index)
            _index1_col_index = _get_col_index(sheet_i, headers, '_index')

            for key, group_indices in g_indices.iteritems():
                if _parent_index in group_indices.keys():
                    for group_row in range(len(group_indices[_parent_index])):
                        group_sheetname_el = ET.SubElement(root, group_sheetname)

                        multi_selects = {}

                        for group_col in range(0, _index1_col_index):
                            header = group_sheet.cell_value(0, group_col)
                            idx = group_indices[_parent_index][group_row]
                            text = group_sheet.cell_value(idx,group_col)
                            if "/" in header:
                                _parse_multi_select_data(multi_selects, header, text)
                                #_handle_multiselects(sheet_i, header, text)
                            else:
                                column_el = ET.SubElement(group_sheetname_el,header)
                                column_el.text = text

                        _gen_multi_selects(group_sheetname_el, multi_selects)

def gen_xml(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)
    headers = _gen_headers(book)
    print headers

    # Get the first worksheet and column names
    data_sheet0 = book.sheet_by_index(0)

    # Loop through rows in first worksheet to output XML, including repeating groups
    for row in range(1, data_sheet0.nrows):
        elems = _gen_xml_elements0(book, headers, row)
        root = elems[row]['root']

        if _has_group(book):
            _gen_group_detail(book, row, headers, data_sheet0, root)

        # Create __version__ element
        version_el = ET.SubElement(root,"__version__")
        version_col_index = _get_col_index(0,headers, '__version__')
        version_el.text = str(data_sheet0.cell_value(row,version_col_index))

        # Create meta element with nested instanceID
        meta_el = ET.SubElement(root,"meta")
        instance_ID_el = ET.SubElement(meta_el, "instanceID")
        _uuid_col_index = _get_col_index(0, headers, '_uuid')
        iID = data_sheet0.cell_value(row, _uuid_col_index)
        iID = iID if len(iID) > 0 else str(uuid.uuid4())
        output_iID = "uuid:"+iID
        instance_ID_el.text = output_iID

        # Create the xml files
        tree = ET.ElementTree(root)
        output_fn = iID + '.xml'
        tree.write(output_fn, pretty_print=True, xml_declaration=True, encoding="utf-8")

if __name__ == "__main__":
    INPUT_EXCEL_FILE = sys.argv[1] # "xls-to-xml-test.xlsx"
    KPI_UID = sys.argv[2]
    KC_UUID = sys.argv[3]

    gen_xml(INPUT_EXCEL_FILE)
