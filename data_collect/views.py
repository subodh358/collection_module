from django.shortcuts import render
import xlrd
import openpyxl
from data_collect.models import Book
from django.http import HttpResponse
# Create your views here.


def upload_file(request):
    if request.method == "POST":
        excel_file = request.FILES["excel_file"]
        wb = openpyxl.load_workbook(excel_file, data_only=True)
        worksheet = wb["Sheet1"]
        add_list = []
        iterrows = iter(worksheet.rows)
        header_objects = next(iterrows)
        for row in iterrows:
            row_data = []
            for cell in row:
                row_data.append(cell.value)
            obj = Book(
                ofp_plan_date=row_data[0],
                f_in_date=row_data[1],
                f_out_date=row_data[2],
                invoice=row_data[3],
                t_qty=row_data[4],
                a_qty=row_data[5],
                a_part=row_data[6],
                rfid_seal=row_data[7],
                container_no=row_data[8],
                seal_no=row_data[9],
                lr_no=row_data[10],
                vehicle_no=row_data[11],
                s_line=row_data[12],
                clearance=row_data[13],
                destination=row_data[14],
                port=row_data[15],
                f_dest=row_data[16],
                product=row_data[17],
                booking_no=row_data[18],
                transporter=row_data[19],
                cont_size=row_data[20],
                detn_days=row_data[21],
                detain_reason=row_data[22],
                mobile_no=row_data[23],
                remarks=row_data[24]
            )
            add_list.append(obj)
        Book.objects.bulk_create(add_list)
        return render(request, 'upload.html')
    else:
        return HttpResponse("Error")
