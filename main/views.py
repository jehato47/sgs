from django.shortcuts import render
from rest_framework.response import Response
from rest_framework.decorators import api_view, permission_classes
from rest_framework.permissions import IsAuthenticated
import docx
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx2pdf import convert
from pdf2image import convert_from_path
from docx.enum.text import WD_COLOR_INDEX
from .serializer import SgsSerializer
from openpyxl import load_workbook
# import os
import img2pdf
import os


@api_view(["GET"])
def createSgs(request):
    import pythoncom
    pythoncom.CoInitialize()
    data = request.data
    data = data.dict()

    wb = load_workbook("C:/Users/LENOVO/Desktop/cemal.xlsx")
    ws = wb.active
    ws["B5"] = data["carBrand"]
    ws["D5"] = data["carModel"]
    ws["E5"] = int(data["carYear"])
    ws["F5"] = int(data["old"])
    ws["G5"] = data["country"]
    ws["H5"] = data["vinNumber"]

    wb.save(f"{data['vinNumber']}.xlsx")

    document = docx.Document("C:/Users/LENOVO/Desktop/2nd.docx")
    table = document.tables[0]

    # company name
    table.column_cells(13)[7].text = data["exporterCompany"]

    # adress
    table.column_cells(13)[8].text = data["exporterAddress"]

    # contact person
    table.column_cells(13)[9].text = data["contactPerson"]

    # email
    table.column_cells(13)[10].text = data["email"]

    # telephone no
    table.column_cells(13)[11].text = data["phone"]

    # Importer ***
    # Company Name
    table.column_cells(14)[7].text = data["importerCompany"]

    # Address
    table.column_cells(14)[8].text = data["importCompanyAddress"]

    # Invoice no / date
    table.column_cells(20)[13].text = data["invoiceNoDate"]
    table.column_cells(20)[13].paragraphs[0].runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW

    # shading_elm_1 = parse_xml(r'<w:shd {} w:fill="faff00"/>'.format(nsdecls('w')))
    # table.column_cells(20)[13]._tc.get_or_add_tcPr().append(shading_elm_1)

    document.tables[0] = table
    document.save("doc.docx")

    convert("doc.docx")
    images = convert_from_path("doc.pdf")
    images[0].resize((700, 600))
    images[0].save("first.jpeg", "jpeg")

    with open(f"{data['vinNumber']}.pdf", "wb") as f:
        f.write(
            img2pdf.convert([i for i in os.listdir(os.curdir) if i.endswith(".jpeg")]))

    os.remove("doc.docx")
    os.remove("doc.pdf")
    os.remove("first.jpeg")

    return Response(True)
