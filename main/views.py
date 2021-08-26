from django.shortcuts import render
from rest_framework.response import Response
from rest_framework.decorators import api_view, permission_classes
from django.core.files.storage import default_storage
from django.core.files.base import ContentFile
import shutil
from rest_framework.permissions import IsAuthenticated
import docx
from docx2pdf import convert
from pdf2image import convert_from_path
from docx.enum.text import WD_COLOR_INDEX
from .serializer import SgsSerializer
from .models import Sgs
from openpyxl import load_workbook
import img2pdf
import os
import convertapi
import djangoProject1.firebasestr as pyr

os.chdir("media/")


#
# @api_view(["GET", "POST"])
# def createSgs(request):
#     import pythoncom
#     pythoncom.CoInitialize()
#     data = request.data
#     data = data.dict()
#
#     print(data)
#     # return Response(True)
#     wb = load_workbook("cemal.xlsx")
#     ws = wb.active
#     ws["B5"] = data["carBrand"]
#     ws["D5"] = data["carModel"]
#     ws["E5"] = int(data["carYear"])
#     ws["F5"] = int(data["old"])
#     ws["G5"] = data["country"]
#     ws["H5"] = data["vinNumber"]
#
#     wb.save("{}.xlsx".format(data['vinNumber']))
#     pyr.upload("{}/{}.xlsx".format(data['vinNumber'], data['vinNumber']), "{}.xlsx".format(data['vinNumber']))
#
#     document = docx.Document("2nd.docx")
#     table = document.tables[0]
#
#     # company name
#     table.column_cells(13)[7].text = data["exporterCompany"]
#
#     # adress
#     table.column_cells(13)[8].text = data["exporterAddress"]
#
#     # contact person
#     table.column_cells(13)[9].text = data["contactPerson"]
#
#     # email
#     table.column_cells(13)[10].text = data["email"]
#
#     # telephone no
#     table.column_cells(13)[11].text = data["phone"]
#
#     # Importer ***
#     # Company Name
#     table.column_cells(14)[7].text = data["importerCompany"]
#
#     # Address
#     table.column_cells(14)[8].text = data["importerAddress"]
#
#     # Invoice no / date
#     table.column_cells(20)[13].text = data["invoiceNoDate"]
#     table.column_cells(20)[13].paragraphs[0].runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW
#
#     # shading_elm_1 = parse_xml(r'<w:shd {} w:fill="faff00"/>'.format(nsdecls('w')))
#     # table.column_cells(20)[13]._tc.get_or_add_tcPr().append(shading_elm_1)
#
#     document.tables[0] = table
#     document.save("doc.docx")
#
#     convertapi.api_secret = 'gKDNm1UdZ94tL5zI'
#     files = convertapi.convert('png', {
#         'File': 'doc.docx',
#     }, from_format='docx').save_files('.')
#     my_file = 'doc.png'
#     base = os.path.splitext(my_file)[0]
#     os.rename(my_file, base + '.jpeg')
#
#     with open("{}.pdf".format(data['vinNumber']), "wb") as f:
#         f.write(
#             img2pdf.convert([i for i in os.listdir(os.curdir) if i.endswith(".jpeg")]))
#
#     sgs = Sgs(pdfFile="{}.pdf".format(data['vinNumber']), xlsxFile="{}.xlsx".format(data['vinNumber']))
#     serializer = SgsSerializer(sgs)
#     if serializer.is_valid():
#         serializer.save()
#     print(1232)
#     print(serializer.errors)
#
#     os.remove("doc.docx")
#     os.remove("doc.jpeg")
#     os.remove("doc-2.png")
#
#     pyr.upload("{}/{}.pdf".format(data['vinNumber'], data['vinNumber']), "{}.pdf".format(data['vinNumber']))
#
#     os.remove("{}.pdf".format(data['vinNumber']))
#     os.remove("{}.xlsx".format(data['vinNumber']))
#
#     Sgs()
#     SgsSerializer()
#
#     return Response(True)


@api_view(["POST"])
def createXlsx(request):
    import pythoncom
    pythoncom.CoInitialize()
    data = request.data
    data = data.dict()

    print(data)
    # return Response(True)
    wb = load_workbook("cemal.xlsx")
    ws = wb.active
    ws["B5"] = data["carBrand"]
    ws["D5"] = data["carModel"]
    ws["E5"] = int(data["carYear"])
    ws["F5"] = int(data["old"])
    ws["G5"] = data["country"]
    ws["H5"] = data["vinNumber"]

    wb.save("{}.xlsx".format(data['vinNumber']))
    pyr.upload("{}/{}.xlsx".format(data['vinNumber'], data['vinNumber']), "{}.xlsx".format(data['vinNumber']))
    os.remove("{}.xlsx".format(data['vinNumber']))

    return Response(True)


@api_view(["POST"])
def createSgs(request):
    data = request.data.dict()
    file = data['invoice']

    filename = file.name.split(".")
    filename[-1] = "jpeg"
    filename = ".".join(filename)

    return Response(True)
    path = default_storage.save(filename, ContentFile(file.read()))

    # shutil.move(path, "files/{}".format(filename))

    import pythoncom
    pythoncom.CoInitialize()
    data = request.data
    data = data.dict()
    document = docx.Document("2nd.docx")
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
    table.column_cells(14)[8].text = data["importerAddress"]

    # Invoice no / date
    table.column_cells(20)[13].text = data["invoiceNoDate"]
    table.column_cells(20)[13].paragraphs[0].runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW

    # shading_elm_1 = parse_xml(r'<w:shd {} w:fill="faff00"/>'.format(nsdecls('w')))
    # table.column_cells(20)[13]._tc.get_or_add_tcPr().append(shading_elm_1)

    document.tables[0] = table
    document.save("doc.docx")

    convertapi.api_secret = 'gKDNm1UdZ94tL5zI'  # eski
    # convertapi.api_secret = '9R7pyBfJESrq7tJ1'  # yeni
    files = convertapi.convert('png', {
        'File': 'doc.docx',
    }, from_format='docx').save_files('.')
    my_file = 'doc.png'
    base = os.path.splitext(my_file)[0]
    os.rename(my_file, base + '.jpeg')

    with open("{}.pdf".format(data['vinNumber']), "wb") as f:
        f.write(
            img2pdf.convert([i for i in os.listdir(os.curdir) if i.endswith(".jpeg")]))

    os.remove("doc.docx")
    os.remove("doc.jpeg")
    os.remove("doc-2.png")

    pyr.upload("{}/{}.pdf".format(data['vinNumber'], data['vinNumber']), "{}.pdf".format(data['vinNumber']))

    os.remove("{}.pdf".format(data['vinNumber']))
    os.remove(filename)

    return Response(True)


@api_view(["GET"])
def sample(request):
    return Response({"success": True})
