import xmltodict, os
import pandas as pd
from flask import *
from flask import send_file


class ProcessVoucharTransactionn():
    df = None
    def processTransaction(self, XML_dict):

        finalData = {
                     "Date" : [],
                     "Transaction Type" : [],
                     "Vch No" : [],
                     "Ref No" : [],
                     "Ref Type" : [],
                     "Ref Date" : [],
                     "Debtor" : [],
                     "Ref Amount" : [],
                     "Amount" : [],
                     "Particulars" : [], 
                     "Vch Type" : [],
                     "Amount Verified" : []
                     }
        
        voucher_list = [tallyMessage['VOUCHER'] for tallyMessage in XML_dict['ENVELOPE']['BODY']['IMPORTDATA']['REQUESTDATA']['TALLYMESSAGE'] if tallyMessage['VOUCHER']['VOUCHERTYPENAME'] == "Receipt"]

        for voucher in voucher_list :

            childTransaction = voucher["ALLLEDGERENTRIES.LIST"][0]["BILLALLOCATIONS.LIST"]
            childTransaction = [childTransaction] if isinstance(childTransaction, dict) else childTransaction

            for index in range(len(childTransaction)):
                finalData["Date"].append(voucher['DATE'])
                finalData["Transaction Type"].append("Child")
                finalData["Vch No"].append(voucher['VOUCHERNUMBER'])
                finalData["Ref No"].append(childTransaction[index]["NAME"])
                finalData["Ref Type"].append(childTransaction[index]["BILLTYPE"])
                finalData["Ref Date"].append("")
                finalData["Debtor"].append(voucher['ALLLEDGERENTRIES.LIST'][0]['LEDGERNAME'])
                finalData["Ref Amount"].append(childTransaction[index]["AMOUNT"])
                finalData["Amount"].append("NA")
                finalData["Particulars"].append(voucher['ALLLEDGERENTRIES.LIST'][0]['LEDGERNAME'])
                finalData["Vch Type"].append(voucher['VOUCHERTYPENAME'])
                finalData["Amount Verified"].append("NA")

            finalData["Date"].append(voucher['DATE'])
            finalData["Transaction Type"].append("Parent")
            finalData["Vch No"].append(voucher['VOUCHERNUMBER'])
            finalData["Ref No"].append("NA")
            finalData["Ref Type"].append("NA")
            finalData["Ref Date"].append("NA")
            finalData["Debtor"].append(voucher['PARTYLEDGERNAME'])
            finalData["Ref Amount"].append("NA")
            finalData["Amount"].append(voucher['ALLLEDGERENTRIES.LIST'][0]['AMOUNT'])
            finalData["Particulars"].append(voucher['PARTYLEDGERNAME'])
            finalData["Vch Type"].append(voucher['VOUCHERTYPENAME'])
            finalData["Amount Verified"].append("NA")

            finalData["Date"].append(voucher['DATE'])
            finalData["Transaction Type"].append("Other")
            finalData["Vch No"].append(voucher['VOUCHERNUMBER'])
            finalData["Ref No"].append("NA")
            finalData["Ref Type"].append("NA")
            finalData["Ref Date"].append("NA")
            finalData["Debtor"].append(voucher['ALLLEDGERENTRIES.LIST'][1]['LEDGERNAME'])
            finalData["Ref Amount"].append("NA")
            finalData["Amount"].append(voucher['ALLLEDGERENTRIES.LIST'][1]['AMOUNT'])
            finalData["Particulars"].append(voucher['ALLLEDGERENTRIES.LIST'][1]['LEDGERNAME'])
            finalData["Vch Type"].append(voucher['VOUCHERTYPENAME'])
            finalData["Amount Verified"].append("NA")

        self.df = pd.DataFrame(finalData)
        self.df['Ref Amount'] = pd.to_numeric(self.df['Ref Amount'], errors='coerce')
        self.df['Amount'] = pd.to_numeric(self.df['Amount'], errors='coerce')

        
        self.df["Amount Verified"] = self.df.apply(self.checkAmountVerified, axis=1)

        self.df['Ref Amount'] = self.df['Ref Amount'].fillna('NA').astype(str)
        self.df['Amount'] = self.df['Amount'].fillna('NA').astype(str)

        self.df["Date"] = pd.to_datetime(self.df['Date'], format='%Y%m%d').dt.strftime("%d-%m-%Y")
        path = os.getcwd() + "\Response.xlsx"
        self.df.to_excel(path)
        print("path : ",path)
        if os.path.exists(path):
            return path
        

    def checkAmountVerified(self, row):
        child_sums = self.df[self.df["Transaction Type"] == "Child"].groupby("Vch No")["Ref Amount"].sum()
        
        if row["Transaction Type"] == "Parent":
            child_sum = child_sums.get(row["Vch No"], 0)
            if child_sum == row["Amount"]:
                return "Yes"
            else:
                return "No"
        return "NA" 


app = Flask(__name__)

@app.route('/upload/', methods=['POST'])
def upload():
    
    if request.method == 'POST':
        
        uploaded_file = request.files["XMLfile"]                    
        if uploaded_file:
            my_xml = uploaded_file.read().decode('utf-8')
            
            XML_dict = xmltodict.parse(my_xml)
            pvt = ProcessVoucharTransactionn()
            file_path = pvt.processTransaction(XML_dict)
 
            return send_file(file_path, as_attachment=True)
    

if __name__ == '__main__':
    app.run(debug=True)
