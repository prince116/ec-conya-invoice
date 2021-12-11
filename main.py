from datetime import datetime as dt
from pandas.io.formats.excel import ExcelFormatter
import os
import pandas as pd

ExcelFormatter.header_style = None
outputFilename: str = "output_" + dt.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"


def processFile(filepath: str):
    df = pd.read_excel(filepath, header=None)
    pd_data = df[1:len(df)]

    dataList: list = []
    header: list = [
        "Date",
        "Patient No.",
        "Patient Name",
        "Item Code",
        "Item Name",
        "Payment Shop",
        "Cash",
        "Cheque",
        "HS-Visa",
        "HS-銀聯",
        "Conya-Contra",
        "Conya-MV",
        "Total Sales",
        "Dentist",
        "Dentist Code"
    ]

    for row, col in pd_data.iterrows():
        try:
            invoiceDate: str = dt.strptime(str(col[0]), "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%d")
        except ValueError:
            invoiceDate: str = "wrong_format"

        patientNo: str = col[1]
        patientName: str = col[2]
        itemCode: str = col[3]
        itemName: str = col[4]
        paymentShop: str = "CTWV" if col[5] == "CTWN" else col[5]
        cash: str = col[6]
        cheque: str = col[7]
        hsVisa: str = col[8]
        hsUnionPay: str = col[9]
        conyaContra: str = col[10]
        conyaMV: str = col[11]
        totalSales: str = col[12]
        dentist: str = col[13]
        dentistCode: str = col[14]

        invoiceData: list = [invoiceDate, patientNo, patientName, itemCode, itemName, paymentShop, cash, cheque, hsVisa, hsUnionPay, conyaContra, conyaMV, totalSales, dentist, dentistCode]
        dataList.append(invoiceData)

        print("[%s/%s]" % (row, len(pd_data)), end='\r')

    excelDf = pd.DataFrame(dataList)
    with pd.ExcelWriter(outputFilename) as writer:
        print("Writing data to the Excel file ...")
        excelDf.to_excel(writer, sheet_name="Sheet1", index=False, header=header)

        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]

        # Set cell format
        cellFormat = workbook.add_format()
        cellFormat.set_align('left')
        worksheet.set_column('A:O', None, cellFormat)

        # Set header format
        headerFormat = workbook.add_format()
        headerFormat.set_bold()
        headerFormat.set_align("left")
        headerFormat.set_font_size(10)
        headerFormat.set_font_name("Calibri Light")

        for index, col in enumerate(header):
            worksheet.write(0, index, col, headerFormat)

        writer.save()
        print("Finished!")


def askForFileName():
    fileName: str = input("Enter File Name: ")
    filePath: str = os.path.join(os.path.abspath(os.getcwd()), fileName)

    if not os.path.exists(filePath):
        print("File dost not exist.")
        return

    processFile(filePath)


def main():
    askForFileName()


if __name__ == "__main__":
    main()
