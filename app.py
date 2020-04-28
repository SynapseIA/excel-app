from flask import Flask, url_for, render_template, request, session
import xlrd
import xlwt
import azure
from azure.storage.blob import BlockBlobService
from xlutils.copy import copy

app=Flask(__name__)

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/_pushData/', methods=['POST'])
def _pushData():
    input1=request.form.get('input1') or None
    input2=request.form.get('input2') or None
    input3=request.form.get('input3') or None
    input4=request.form.get('input4') or None

    if not input1 or not input2 or not input3 or not input4:
        return "Please provide all data <a href='/'> Push data <a>"
    else: 
        operation(input1, input2, input3, input4)
    return "done"


def operation(input1, input2, input3, input4):
    azure_storage_account_name = "synapsestac"
    azure_storage_account_key2 = "Gfmxsc3TvQYynAiVB4uFGEnfuUFYT2KAEKarWcvvdBtCDU+2nBcSPjjVYtYgbjjr1xoFuXsjeWbj7+uoiKHhYA=="
    
    Blob_Service = BlockBlobService(azure_storage_account_name,azure_storage_account_key2)

    blobs = Blob_Service.list_blobs("excel-app")

    for blob in blobs :
        blob_name = blob.name
        if blob_name == "excel-app.xlsx" : 
            a = blob_name

    Blob_Service.get_blob_to_path("excel-app", a, "data_xl")
    data = xlrd.open_workbook("data_xl")
    wb = copy(data)
    Sheet1=wb.get_sheet(0)
    Sheet1.write(1,2,int(input1))
    Sheet1.write(2,2,int(input2))
    Sheet1.write(3,2,int(input3))
    Sheet1.write(4,2,int(input4))
    wb.save('excel-app.xlsx')

if __name__=="__main__":
    app.run(debug=True)
