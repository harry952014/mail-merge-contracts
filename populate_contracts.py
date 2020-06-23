from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import pandas as pd
import xlrd
from docx2pdf import convert
import convert2pdf

template = "Catlin Solar - Community Solar Subscriber Agreement (clean).docx"
# document = MailMerge(template)
# print(document.get_merge_fields())


#xl_file = pd.read_excel('Bulkproject - Catlin ALTICE  2020 05 28.xlsx', sheet_name="New Customers")
xl_file = pd.read_excel('Altice NYSEG Meters Not in Accounting Template.xlsx', sheet_name="Meters Not in Acct Template");

# book1 = xlrd.open_workbook('Bulkproject - Catlin ALTICE  2020 05 28.xlsx')
# sheet1 = book1.sheet_by_name("New Customers");

word_contract_folder = '/Users/haris/Documents/PowerMarket Docs/Green Street/Catlin/Contracts Work/Altice Contracts (Populated)/Word/';
pdf_contract_folder = '/Users/haris/Documents/PowerMarket Docs/Green Street/Catlin/Contracts Work/Altice Contracts (Populated)/PDF/';

for ind,row in xl_file.iterrows():
    # print(str(row['Zip_1']))
    street = str(row['Address'])
    city = str(row['City'])
    zip = str(row['Zip'])
    alloc = str('{0:.3%}'.format(row['% Allocation']))
    kw = str(row['System size (kW)'])

    print(street, " " , city , " " , zip , " " , alloc , " " , kw)

    document = MailMerge(template)
    document.merge(
        Street=street,
        City=city,
        Zip=zip,
        allocation=alloc,
        sysSize=kw);

    word_contract_name = '/Users/haris/Documents/PowerMarket Docs/Green Street/Catlin/Contracts Work/Altice Contracts (Populated)/Word/Catlin Solar 1 LLC GSPP Subscriber Agreement (Altice ' + str(row['Utility Account']) + ').docx'
    pdf_contract_name = '/Users/haris/Documents/PowerMarket Docs/Green Street/Catlin/Contracts Work/Altice Contracts (Populated)/PDF/Catlin Solar 1 LLC GSPP Subscriber Agreement (Altice ' + str(row['Utility Account']) + ').pdf'

    document.write(word_contract_name)
    convert(word_contract_name, pdf_contract_name)



# convert(word_contract_folder, pdf_contract_folder)



