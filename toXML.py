# Экспорт данных из Ecxel в XML
# Вход: xlsx_name файл
# Выход: outxml_name файл
################################

import xml.etree.cElementTree as ET
from openpyxl import load_workbook


def indent(elem, level=0):
    i = "\n" + level * "\t"
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "\t"
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level + 1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i


def main():
    xlsx_name = "in_excel.xlsx"
    outxml_name = "out_xml.xml"

    # загрузка файла xlsx
    wb = load_workbook(xlsx_name)
#   ws = wb.worksheets[0]
    sheet = wb.active  # Активный лист в книге

    # формирование xml файла
    root = ET.Element("EconomProfleGraf")
    for dt, okato, pok_id, val in zip(
        sheet["A"][1:], sheet["B"][1:], sheet["C"][1:], sheet["D"][1:]
    ):
        ET.SubElement(
            root,
            'row DT="'
            + dt.value
            + '" OKATO="'
            + str(okato.value)
            + '" POK_ID="'
            + str(pok_id.value)
            + '" VAL="'
            + str(val.value)
            + '"',)

    indent(root)
    out = open(outxml_name, "wb")
    out.write(b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
    ET.ElementTree(root).write(out, encoding="UTF-8", xml_declaration=False)
    print(f'Экспорт файла {outxml_name} завершен!')


if __name__ == "__main__":
    main()
