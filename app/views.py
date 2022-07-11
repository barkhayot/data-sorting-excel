from django.shortcuts import render
import pandas as pd
from django.http import HttpResponse, FileResponse
import os
from time import time
from django.http import HttpResponseRedirect
from django.views import View
from django.core.mail import EmailMessage
from django.conf import settings
from .forms import EmailForm
from zipfile import ZipFile

import pandas as pd
from time import time


def generate(file_excel, sheet_name, folder, name):
    data = pd.read_excel(file_excel, sheet_name=sheet_name)
    sheet = {}

    for col in data:
        col_list = list(data[col])
        column = []
        run = 0
        for items in col_list:
            splited = str(data["배송메모2"][run]).split(" /")
            if (col == "배송메시지"):
                added = 0
                for i in splited:
                    if ("택배배송" in str(i)) or ("새벽배송" in str(i)):
                        column.append(i)
                        added += 1
                if added == 0:
                    column.append("택배배송")
            
                sheet[col] = column
            elif (col == "기타10"):
                added = 0
                for i in splited:
                    if ("출입방법:" in str(i)):
                        i = i.replace("  출입방법:", "출입방법:")
                        column.append(i)
                        added += 1
                if added == 0:
                    column.append("")
            
                sheet[col] = column
            elif (col == "기타9"):
                added = 0
                for i in splited:
                    if ("문자알림:" in str(i)):
                        if ("수신거부" in str(i)):
                            column.append("9")
                        elif ("배송완료 후 즉시" in str(i)):
                            column.append("2")
                        else:
                            column.append("1")
                        added += 1
                if added == 0:
                    column.append("1")
            
                sheet[col] = column
            elif (col == "배송메모"):
                added = 0
                for i in splited:
                    if ("출입방법:" not in str(i)) and ("문자알림:" not in str(i)) and ("택배배송" not in str(i)) and ("새벽배송" not in str(i)):
                        if len(i)>0:
                            if i[:3]=="   ":
                                column.append(i[3:])
                            elif i[:2]=="  ":
                                column.append(i[2:])
                            elif i[0]==" ":
                                column.append(i[1:])
                            elif "nan" in str(i):
                                column.append("")
                            else:
                                column.append(i)
                        else:
                            column.append("")
                            
                        added += 1
                if added == 0 or len(splited)==0:
                    column.append("")
            
                sheet[col] = column
            else:
                sheet[col] = list(data[col])
            run+=1


    writer = pd.ExcelWriter(f'{folder}/filtered_{name}.xlsx')
    sheet_name = "filtered"
    sheet = pd.DataFrame(sheet)
    sheet.to_excel(writer, sheet_name=sheet_name, index=False)
    writer.save()

    data = pd.read_excel(f"{folder}/filtered_{name}.xlsx", sheet_name=sheet_name)

    sheets = ["음성_저온_택배", "음성_상온_택배", "음성_저온_수도권새벽배송", "음성_상온_수도권새벽배송", "음성_상온_영남새벽배송", 
            "음성_저온_영남새벽배송", "영천_수도권새벽배송", "영천_영남새벽배송", "영천_택배배송"
            ]
    supplier_name = "공급사명"
    order = "배송메시지"

    work_book = {}
    for sheet in sheets:
        
        work_book[sheet] = {}

        if True:        
            for col in data:
                column = []
                run = 0
                for i in data[col]:
                    if (sheet == "음성_저온_택배") and ("택배배송" in data[order][run]) and ("음성실재고(저온)" in data[supplier_name][run]):
                        if str(i) == "nan":
                            column.append("")
                        else:
                            column.append(str(i))
                    
                    elif (sheet == "음성_상온_택배") and ("택배배송" in data[order][run]) and ("음성실재고(상온)" in data[supplier_name][run]):
                        if str(i) == "nan":
                            column.append("")
                        else:
                            column.append(str(i))
                    
                    elif (sheet == "음성_저온_수도권새벽배송") and ("(서울/경기/인천) 새벽배송" in data[order][run]) and ("음성실재고(저온)" in data[supplier_name][run]):
                        if str(i) == "nan":
                            column.append("")
                        else:
                            column.append(str(i))
                    
                    elif (sheet == "음성_상온_수도권새벽배송") and ("(서울/경기/인천) 새벽배송" in data[order][run]) and ("음성실재고(상온)" in data[supplier_name][run]):
                        if str(i) == "nan":
                            column.append("")
                        else:
                            column.append(str(i))
                    
                    elif (sheet == "음성_상온_영남새벽배송") and (("(부산/경남) 새벽배송" in data[order][run]) or ("(대구) 새벽배송" in data[order][run])) and ("음성실재고(상온)" in data[supplier_name][run]):
                        if str(i) == "nan":
                            column.append("")
                        else:
                            column.append(str(i))
                    
                    elif (sheet == "음성_저온_영남새벽배송") and (("(부산/경남) 새벽배송" in data[order][run]) or ("(대구) 새벽배송" in data[order][run])) and ("음성실재고(저온)" in data[supplier_name][run]):
                        if str(i) == "nan":
                            column.append("")
                        else:
                            column.append(str(i))
                            
                    elif (sheet == "영천_수도권새벽배송") and ("(서울/경기/인천) 새벽배송" in data[order][run]) and ("데이웰즈(냉장/냉동)" in data[supplier_name][run]):
                        if str(i) == "nan":
                            column.append("")
                        else:
                            column.append(str(i))
                    
                    elif (sheet == "영천_영남새벽배송") and (("(부산/경남) 새벽배송" in data[order][run]) or ("(대구) 새벽배송" in data[order][run])) and ("데이웰즈(냉장/냉동)" in data[supplier_name][run]):
                        if str(i) == "nan":
                            column.append("")
                        else:
                            column.append(str(i))
                    elif (sheet == "영천_택배배송") and ("택배배송" in data[order][run]) and ("데이웰즈(냉장/냉동)" in data[supplier_name][run]):
                        if str(i) == "nan":
                            column.append("")
                        else:
                            column.append(str(i))
                    
                    run+=1
                work_book[sheet][col] = column

    writer = pd.ExcelWriter(f'{folder}/worker_{name}.xlsx')
    data.to_excel(writer, sheet_name=sheet_name, index=False)
    for sheet_name in work_book:
            data = pd.DataFrame(work_book[sheet_name])
            data.to_excel(writer, sheet_name=sheet_name, index=False)
    writer.save()

    writer = pd.ExcelWriter(f'{folder}/truck_{name}.xlsx')
    for sheet_name in work_book:
        if sheet_name == "음성_상온_영남새벽배송" or sheet_name == "음성_저온_영남새벽배송" or sheet_name == "영천_수도권새벽배송":
            data = pd.DataFrame(work_book[sheet_name])
            data.to_excel(writer, sheet_name=sheet_name, index=False)
    writer.save()

    writer = pd.ExcelWriter(f'{folder}/location_{name}.xlsx')
    for sheet_name in work_book:
        if sheet_name == "음성_상온_영남새벽배송" or sheet_name == "음성_저온_영남새벽배송" or sheet_name == "영천_영남새벽배송":
            data = pd.DataFrame(work_book[sheet_name])
            data.to_excel(writer, sheet_name=sheet_name, index=False)
    writer.save()

    return f'{folder}/filtered_{name}.xlsx', f'{folder}/worker_{name}.xlsx', f'{folder}/truck_{name}.xlsx', f'{folder}/location_{name}.xlsx'



def view(request):
    if request.method=="POST":
        file = request.FILES['myfile']
        data = pd.read_excel(file)
        sheets = []
        for i in data["배송메모2"]:
            #if "(대구)" in str(i):
            res = str(i).split(" /")
            if len(res)>=2:
                sheets.append(res[0])
            else:
                sheets.append("Non")
            
        sheets = list(set(sheets))
        all_data = {}
        for sheet in sheets:
            all_data[sheet] = {}

                
            for col in data:
                run = 0
                
                
                if col == "배송메모2":
                    multi = {"배송메모1":[],
                        "배송메모2":[],
                        "배송메모3":[],
                        "배송메모4":[],
                        }
                    for i in data[col]:
                        if sheet == "Non" and len((str(data["배송메모2"][run]).split(" /")))<2:
                            try:
                                multi["배송메모1"].append(str(data["배송메모2"][run]).split(" /")[0])
                            except:
                                multi["배송메모1"].append(str("Non"))
                            try:
                                multi["배송메모2"].append(str(data["배송메모2"][run]).split(" /")[1])
                            except:
                                multi["배송메모2"].append(str("Non"))
                            try:
                                multi["배송메모3"].append(str(data["배송메모2"][run]).split(" /")[2])
                            except:
                                multi["배송메모3"].append(str("Non"))
                            try:
                                multi["배송메모4"].append(str(data["배송메모2"][run]).split(" /")[3])
                            except:
                                multi["배송메모4"].append(str("Non"))
                            
                        elif sheet in str(data["배송메모2"][run]):
                            try:
                                multi["배송메모1"].append(str(data["배송메모2"][run]).split(" /")[0])
                            except:
                                multi["배송메모1"].append(str("Non"))
                            try:
                                multi["배송메모2"].append(str(data["배송메모2"][run]).split(" /")[1])
                            except:
                                multi["배송메모2"].append(str("Non"))
                            try:
                                multi["배송메모3"].append(str(data["배송메모2"][run]).split(" /")[2])
                            except:
                                multi["배송메모3"].append(str("Non"))
                            try:
                                multi["배송메모4"].append(str(data["배송메모2"][run]).split(" /")[3])
                            except:
                                multi["배송메모4"].append(str("Non"))
                        run+=1
                    all_data[sheet].update(multi)

                else:
                    column = []
                    for i in data[col]:
                        if sheet == "Non" and len((str(data["배송메모2"][run]).split(" /")))<2:
                            column.append(i)
                        elif sheet in str(data["배송메모2"][run]):
                            column.append(i)

                        run+=1

                    all_data[sheet][col] = column
        name = f'output-{str(time()).split(".")[-1]}.xlsx'
        writer = pd.ExcelWriter(name)
        for i in all_data:
            #with pd.ExcelWriter("new2.xlsx", mode="a", engine="openpyxl") as writer:
                data = pd.DataFrame(all_data[i])
                sheet_name=i.replace("/", "-")
                data.to_excel(writer, sheet_name=sheet_name, index=False)
        writer.save()
        # download side
        
        # results = pd.DataFrame()
        response = FileResponse(open(name, 'rb'))
        os.remove(name)
        # response = HttpResponse(,content_type='application/vnd.ms-excel')
        # response['Content-Disposition'] = 'attachment; filename="persons.xls"'
        #results.to_csv(path_or_buf=response,sep=';',float_format='%.2f',index=False,decimal=",")
        return response
    return render(request, 'view.html')


# def send_file_to_email(request):
#     return render(request, 'email.html')



# def send_file_to_email(request):
#     form = EmailForm()
#     if request.method == "POST":
#         form = EmailForm(request.POST, request.FILES)
#         if form.is_valid():
#             subject = form.cleaned_data['subject']
#             message = form.cleaned_data['message']
#             email = form.cleaned_data['email']
#             files = request.FILES.getlist('attach')
#             mail = EmailMessage(subject, message, settings.EMAIL_HOST_USER, [email])
#             print(mail)
#             print(files)
#             for f in files:
#                 mail.attach(f.name, f.read(), f.content_type)
#             mail.send()
#             return render(request, 'email.html', {'email_form': form, 'error_message': 'Sent email to %s'%email})
#             # except:
#             #     return render(request, 'email.html', {'email_form': form, 'error_message': 'Either the attachment is too big or corrupt'})

#         return render(request, 'email.html', {'email_form': form, 'error_message': 'Unable to send email. Please try again later'})
    
#     context = {
#             'form' : form
#         }
#     return render(request, 'email.html', context)

# class EmailAttachementView(View):
#     form_class = EmailForm
#     template_name = 'emailattachment.html'

#     def get(self, request, *args, **kwargs):
#         form = self.form_class()
#         return render(request, self.template_name, {'email_form': form})

#     def post(self, request, *args, **kwargs):
#         form = self.form_class(request.POST, request.FILES)

#         if form.is_valid():
            
#             subject = form.cleaned_data['subject']
#             message = form.cleaned_data['message']
#             email = form.cleaned_data['email']
#             files = request.FILES.getlist('attach')

#             try:
#                 mail = EmailMessage(subject, message, settings.EMAIL_HOST_USER, [email])
#                 for f in files:
#                     mail.attach(f.name, f.read(), f.content_type)
#                 mail.send()
#                 return render(request, self.template_name, {'email_form': form, 'error_message': 'Sent email to %s'%email})
#             except:
#                 return render(request, self.template_name, {'email_form': form, 'error_message': 'Either the attachment is too big or corrupt'})

#         return render(request, self.template_name, {'email_form': form, 'error_message': 'Unable to send email. Please try again later'})



# def filter_data(request):
#     if request.method=="POST":
#         file_excel = request.FILES['myfile']
        
#         ### there should be a form to upload file and enter sheet_name
#         sheet_name = "Worksheet" ##entered sheet name equal to this
#         folder = "temp" ## enter temporary folder name
#         name = str(time()).split("-")[-1] #end name

#         first, second, third, fourth = generate(file_excel, sheet_name, folder, name)

                
                
#         first = FileResponse(open(first, 'rb'))
#         second = FileResponse(open(second, 'rb'))
#         third = FileResponse(open(third, 'rb'))
#         fourth = FileResponse(open(fourth, 'rb'))
#         return first, second, third, fourth

#     return render(request, 'filter.html')

def filter_data(request):
    if request.method=="POST":
        file_excel = request.FILES['myfile']
        
        ### there should be a form to upload file and enter sheet_name
        sheet_name = "Worksheet" ##entered sheet name equal to this
        folder = "temp" ## enter temporary folder name
        name = str(time()).split(".")[-1] #end name

        first, second, third, fourth = generate(file_excel, sheet_name, folder, name)
        name = f"temp/outputs_{name}.zip"
        with ZipFile(name, mode="a") as archive:
            archive.write(first)
            os.remove(first)
            archive.write(second)
            os.remove(second)
            archive.write(third)
            os.remove(third)
            archive.write(fourth)
            os.remove(fourth)

        
        response = FileResponse(open(name, "rb"))
        os.remove(name)
        return response

    return render(request, 'filter.html')