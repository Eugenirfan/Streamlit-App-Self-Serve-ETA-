import streamlit as st
import streamlit.components.v1 as components
import os
import win32com.client
os.environ["CUDA_VISIBLE_DEVICES"] = "-1"
import pyodbc
import pandas as pd

conn = pyodbc.connect('Driver={ODBC Driver 17 for SQL Server};'
                      'Server=server;'
                      'Database=Database;'
                      'Trusted_Connection=yes;')
cursor = conn.cursor()


def app():
    buyer_email = pd.read_excel(r'location\.xlsx')
    x = st.text_input("Enter the Item Number: ")

    question = st.text_input(" Do you have a specific Warehouse? Respond Yes or No ")

    if question =='Yes':
        y = st.text_input("Enter the Shipping Warehouse: ")
        query = (
            "SELECT dwh.PurchaseOrderDetails.ItemNumber, dwh.PurchaseOrderDetails.BranchPlantCode, dwh.PurchaseOrderDetails.StockingStatusCodeBranchPlant, dwh.PurchaseOrderDetails.OrderNumber, dwh.PurchaseOrderDetails.QtyOpenInPurchasingUOM, dwh.PurchaseOrderDetails.PurhasingUOM, dwh.PurchaseOrderDetails.BuyerName, dwh.PurchaseOrderDetails.SupplierName, dwh.PurchaseOrderDetails.IsPOLineComplete, dwh.PurchaseOrderDetails.OrderType, dwh.PurchaseOrderDetails.RequestedDate, dwh.PurchaseOrderDetails.FirstReceiptDate, dwh.PurchaseOrderDetails.ABC1, dwh.PurchaseOrderDetails.ABC3, dwh.InventorySnapshotItemBranch.QtyOnHandInPurchasingUOM, dwh.InventorySnapshotItemBranch.QtyAvailableInPurchasingUOM\n"
            "FROM dwh.PurchaseOrderDetails LEFT JOIN dwh.InventorySnapshotItemBranch ON (dwh.PurchaseOrderDetails.BranchPlantCode = dwh.InventorySnapshotItemBranch.BranchPlantCode) AND (dwh.PurchaseOrderDetails.ItemNumber = dwh.InventorySnapshotItemBranch.ItemNumber)\n"
            "WHERE dwh.PurchaseOrderDetails.ItemNumber = ? AND dwh.PurchaseOrderDetails.BranchPlantCode = ?\n"
            "GROUP BY dwh.PurchaseOrderDetails.ItemNumber, dwh.PurchaseOrderDetails.BranchPlantCode, dwh.PurchaseOrderDetails.StockingStatusCodeBranchPlant, dwh.PurchaseOrderDetails.OrderNumber, dwh.PurchaseOrderDetails.QtyOpenInPurchasingUOM, dwh.PurchaseOrderDetails.PurhasingUOM, dwh.PurchaseOrderDetails.BuyerName, dwh.PurchaseOrderDetails.SupplierName, dwh.PurchaseOrderDetails.IsPOLineComplete, dwh.PurchaseOrderDetails.OrderType, dwh.PurchaseOrderDetails.RequestedDate, dwh.PurchaseOrderDetails.FirstReceiptDate, dwh.PurchaseOrderDetails.ABC1, dwh.PurchaseOrderDetails.ABC3, dwh.InventorySnapshotItemBranch.QtyOnHandInPurchasingUOM, dwh.InventorySnapshotItemBranch.QtyAvailableInPurchasingUOM\n"
            "HAVING (((dwh.PurchaseOrderDetails.IsPOLineComplete)='N') AND (Not (dwh.PurchaseOrderDetails.OrderType)='OD'))")
        data = pd.read_sql(query, conn, params=(x,y))
        mydata = data.set_index('ItemNumber')
        expander = st.beta_expander("Output")
        expander.dataframe(mydata,width=2000,height=200)
    elif question =='No':
        query = (
            "SELECT dwh.PurchaseOrderDetails.ItemNumber, dwh.PurchaseOrderDetails.BranchPlantCode, dwh.PurchaseOrderDetails.StockingStatusCodeBranchPlant, dwh.PurchaseOrderDetails.OrderNumber, dwh.PurchaseOrderDetails.QtyOpenInPurchasingUOM, dwh.PurchaseOrderDetails.PurhasingUOM, dwh.PurchaseOrderDetails.BuyerName, dwh.PurchaseOrderDetails.SupplierName, dwh.PurchaseOrderDetails.IsPOLineComplete, dwh.PurchaseOrderDetails.OrderType, dwh.PurchaseOrderDetails.RequestedDate, dwh.PurchaseOrderDetails.FirstReceiptDate, dwh.PurchaseOrderDetails.ABC1, dwh.PurchaseOrderDetails.ABC3, dwh.InventorySnapshotItemBranch.QtyOnHandInPurchasingUOM, dwh.InventorySnapshotItemBranch.QtyAvailableInPurchasingUOM\n"
            "FROM dwh.PurchaseOrderDetails LEFT JOIN dwh.InventorySnapshotItemBranch ON (dwh.PurchaseOrderDetails.BranchPlantCode = dwh.InventorySnapshotItemBranch.BranchPlantCode) AND (dwh.PurchaseOrderDetails.ItemNumber = dwh.InventorySnapshotItemBranch.ItemNumber)\n"
            "WHERE dwh.PurchaseOrderDetails.ItemNumber = ?\n"
            "GROUP BY dwh.PurchaseOrderDetails.ItemNumber, dwh.PurchaseOrderDetails.BranchPlantCode, dwh.PurchaseOrderDetails.StockingStatusCodeBranchPlant, dwh.PurchaseOrderDetails.OrderNumber, dwh.PurchaseOrderDetails.QtyOpenInPurchasingUOM, dwh.PurchaseOrderDetails.PurhasingUOM, dwh.PurchaseOrderDetails.BuyerName, dwh.PurchaseOrderDetails.SupplierName, dwh.PurchaseOrderDetails.IsPOLineComplete, dwh.PurchaseOrderDetails.OrderType, dwh.PurchaseOrderDetails.RequestedDate, dwh.PurchaseOrderDetails.FirstReceiptDate, dwh.PurchaseOrderDetails.ABC1, dwh.PurchaseOrderDetails.ABC3, dwh.InventorySnapshotItemBranch.QtyOnHandInPurchasingUOM, dwh.InventorySnapshotItemBranch.QtyAvailableInPurchasingUOM\n"
            "HAVING (((dwh.PurchaseOrderDetails.IsPOLineComplete)='N') AND (Not (dwh.PurchaseOrderDetails.OrderType)='OD'))")
        data = pd.read_sql(query, conn, params={x})
        mydata = data.set_index('ItemNumber')
        expander = st.beta_expander("Output")
        expander.dataframe(mydata, width=2000, height=200)

    ask = st.text_input("Do you need more info? Respond Yes or No")
    if ask == 'Yes':
        email_data = pd.merge(mydata, buyer_email, on=['BuyerName'], how='inner')
        email_analyst = email_data['Email'].iloc[0]

        more_info = st.text_input(" Please explain reason for more Information. \n eg: Outdated eta etc... ")

        name = st.text_input(" Please enter your name and department eg:'Irfan from Supply Chain' ")

        email = st.text_input(" Please enter your work email ")

        html1 = data.to_html()

        subject = "ETA REQUEST"
        outlook = win32com.client.Dispatch('outlook.application')

        mail_send = outlook.CreateItem(0)
        #test_email = 'mohamedirfan.suffeerahmed@email.com'
        mail_send.To = email_analyst
        mail_send.Subject = 'ETA REQUEST ' + name
        body = '</h1>Hello</h1>' + name + ' is requesting eta information on ' + x + '<br><br>'+ 'Reason: ' +more_info + '<br><br>'+ html1 + '<br><br>Their email address is ' + email + '<br><br> Thank you'
        mail_send.HTMLBody = (body)

            # mail_send.HTMLBody =  html1
            # mail_send.CC='mohamedirfan.suffeerahmed@mail.com'
        if st.button('Request'):
            mail_send.Send()
            st.write('Thank you for using supply chain eta request. We will reach out to you shortly :)')
    elif ask == 'No':
        st.write('Thank you for using supply chain eta request. Have a nice day :)')









