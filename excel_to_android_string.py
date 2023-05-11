import argparse
import xml.etree.ElementTree as ET
from openpyxl import load_workbook

parser = argparse.ArgumentParser()
parser.add_argument('-f', '--file', type=str, required=True,
                    help='Excel file path')
parser.add_argument('-s', '--sheet', type=str, default='Sheet1',
                    help='Sheet name in Excel file, default is "Sheet1"')
parser.add_argument('-k', '--key', type=int, default=1,
                    help='Index of key column in Excel file, default is 1')
parser.add_argument('-v', '--value', type=int, default=2,
                    help='Index of value column in Excel file, default is 2')
parser.add_argument('-o', '--output', type=str, default='strings.xml',
                    help='Output file name, default is "strings.xml"')
args = parser.parse_args()

def excel_to_xml(excel_file, sheet_name, key_index, value_index, output_file):
    workbook = load_workbook(filename=excel_file)
    worksheet = workbook[sheet_name]

    root = ET.Element("resources")
    for row in worksheet.values:
        key = str(row[key_index-1])
        value = str(row[value_index-1])
        item = ET.SubElement(root, "string", name=key)
        item.text = value

    tree = ET.ElementTree(root)
    tree.write(output_file, encoding='utf-8')

excel_to_xml(args.file, args.sheet, args.key, args.value, args.output)