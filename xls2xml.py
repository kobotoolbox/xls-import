import xlrd
import sys
import xml.etree.ElementTree as ET

"""
sample command line:
python xls2xml.py xls-to-xml-test.xlsx a123456789098765432123
"""

"""
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
"""

#----------------------------------------------------------------------
def gen_xml(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)

    # get the first worksheet
    data_sheet1 = book.sheet_by_index(0)

    root = ET.Element(kpi_uid)
    elem = ET.SubElement(root, "formhub")
    tree = ET.ElementTree(root)
    print '<?xml version="1.0" ?>'

    for colname in data_sheet1.row_values(0)[:5]:
        colname = ET.SubElement(root, colname)
    ET.SubElement(root,"meta")
    tree = ET.ElementTree(root)

    tree.write(sys.stdout)

#----------------------------------------------------------------------
if __name__ == "__main__":
    input_excel_file = sys.argv[1] # "xls-to-xml-test.xlsx"
    kpi_uid = sys.argv[2]
    gen_xml(input_excel_file)
