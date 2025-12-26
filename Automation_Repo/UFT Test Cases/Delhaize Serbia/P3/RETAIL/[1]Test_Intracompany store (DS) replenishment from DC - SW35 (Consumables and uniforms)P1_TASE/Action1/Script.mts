'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	DataRowSet= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Intracompany store (DS) replenishment from DC - SW35 (Consumables and uniforms)P1
'.................Author : TCS        :Bitan
'................ Creation Date    : 7th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrTestCaseName = "Test_Intracompany store (DS) replenishment from DC - SW35 (Consumables and uniforms)P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DS\RETAIL\DT_Intracompany store (DS) replenishment from DC - SW35 (Consumables and uniforms)p1_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'--------------------------MM43---------------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM43_0100_ARTICLE,False) 
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen Description",DT_MM43_1211_TABLE_CELL,False)
Call PressEnter()
Call ClickButtonIfExist("Back   \(F3\)",False)

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM43_1211_TABLECELL_ARTICLE_1,False) 
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen Description",DT_MM43_1211_TABLE_CELL,False)
Call PressEnter()


'''Call SetTextbox("Article","RMMW1-MATNR","",DT_MM43_0100_ARTICLE,False) 
'''Call TakeScreenShot()
'''Call PressEnter()
'''Wait(1)
'''Call TakeScreenShot()

'''Call ClickButtonIfExist("Back   \(F3\)",False)
'''wait(2)
'''Call SetTextbox("Article","RMMW1-MATNR","",DT_MM43_0100_ARTICLE_OCC1,False)
'''
'''Call TakeScreenShot()
'''Call PressEnter()
'''Wait(1)
'''Call TakeScreenShot()
'''
'''Call ClickButtonIfExist("Back   \(F3\)",False)
'''wait(2)

'----------------------Tcode ME21N----------------------------
Call SetTcode(DT_MM43_4008_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MM43_4008_OKCD)

Call SetComboByKey("MEPO_TOPLINE-BSART",DT_MM43_1105_MEPO_TOPLINEBSART)
Call SetTextbox("Supplying Site","MEPO_TOPLINE-SUPERFIELD","",DT_MM43_1105_SUPPLYING_SITE,False) 
Call TakeScreenShot()
Call SelectTab("HEADER_DETAIL","Additional Data",False)
Call TakeScreenShot()
Call SelectTab("HEADER_DETAIL","Org. Data",False)
Wait(2)

Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_MM43_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_MM43_1221_PURCH_GROUP,False)     
Call PressEnter() 
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_MM43_1221_COMPANY_CODE,False)    
Call PressEnter()

Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_MM43_1211_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_MM43_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_MM43_1211_TABLECELL_SITE_0,False)
Call PressEnter()

Call ClickButtonIfExist("Continue",True)
wait(2)
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_DELIVERY_DATE),False) 
Call PressEnter()
Call ClickButtonIfExist("Continue",True)
wait(2)

Call SetTableData("SAPLMEGUITC_1211","Article","2","","",DT_MM43_1211_TABLECELL_ARTICLE_1,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","2","","",DT_MM43_1211_TABLECELL_PO_QUANTITY_1,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_MM43_1211_TABLECELL_SITE_1,False)
Call PressEnter()

Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",ConvertDate(DT_DELIVERY_DATE),False) 
Call PressEnter()
Call ClickButtonIfExist("Continue",True)
wait(2)

Call TakeScreenShot()

Call SelectMenuBar("Purchase Order;Save")
Wait(2)
Call TakeScreenShot()
Call ClickButtonIfExist("Save",True)
Wait(2)
Call TakeScreenShot()


Call GetStatusBar("item2","DT_STO_NUMBER_OUTPUT")
Call WriteRunTimeDataToExcel("DT_MM43_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR",DT_STO_NUMBER_OUTPUT)
'Call VerifyStatusBar("STO Retail created under the number " & DT_STO_NUMBER_OUTPUT)
Call VerifyStatusBar("STO Retail created under the number " & DT_MM43_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)

''''''''''''Lines of code from Line 128 to Line 167  implemented as part of the resolution provided for the defect 34372"""""""""""""""""""
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_VL10G_0500_OKCD) 
Call PressEnter()  
Call TakeScreenShot
Call SetTextbox("Shipping Point/Receiving Pt","ST_VSTEL-LOW","",DT_VL10G_1000_SHIPPING_POINTRECEIVING_PT,False)
Call SetTextbox("Deliv\. Creation Date","ST_LEDAT-LOW","",ConvertDate(DT_VL10G_1000_DELIV_CREATION_DATE),False)
Call SetTextbox("to","ST_LEDAT-HIGH","",ConvertDate(DT_VL10G_1000_TO),False)
Call SetTextbox("CalcRuleDefltDlvCrDt","P_LERUL","","",False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP_ORDER_CRITERIA","Sales Orders",False)
Call TakeScreenShot
Call SetTextboxNoLabel("ST_VBELN-LOW","",DT_VL10G_1020_SALES_DOCUMENT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP_ORDER_CRITERIA","Purchase Orders",False)
Call TakeScreenShot
Call SetTextbox("Purchasing Document","ST_EBELN-LOW","",DT_VL10G_1030_PURCHASING_DOCUMENT,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call FindRowNumber("", "Originating Document", DT_VL10G_1030_PURCHASING_DOCUMENT, "DT_RowNumber")
wait(2)
Call SelectRowGuiGridbyRowNo("",0,DT_RowNumber,False)
Call TakeScreenShot
Call ClickButtonIfExist("Create Delivery in Background   \(Shift\+F7\)",False)
Call TakeScreenShot

Call GetGridContent("",0,"SD Document",2,"Traffic light","S_TL_G","DT_VL10G_0500_OUTBOUND_DELIVERY_OUTPUT")
' WriteRunTimeDataToExcel(strVariableName, strVariableValue)
Call WriteRunTimeDataToExcel("DT_VL10G_0500_OUTBOUND_DELIVERY",DT_VL10G_0500_OUTBOUND_DELIVERY_OUTPUT)

''''''''''''''''''''''''''''''''''''''''''''''''''''''Log Off''''''''''''''''''
Call LogOff()
Call FinalStatus ()

