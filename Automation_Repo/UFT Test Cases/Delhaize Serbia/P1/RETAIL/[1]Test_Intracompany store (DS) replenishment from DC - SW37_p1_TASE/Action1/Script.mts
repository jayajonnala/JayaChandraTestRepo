
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Intracompany store (DS)  - SW37_p1
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


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
'.................Test Script Name : Test_Intracompany store (DS)  - SW37_p1
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 11th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Intracompany store (DS) - SW37_p1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Intracompany store (DS) replenishment from DC - SW31_p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode ME21N----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Enter the Details
Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
Call SetTextbox("Supplying Site","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_SUPPLYING_SITE,False) 
Call TakeScreenShot()
Call SelectTab("HEADER_DETAIL","Additional Data",False)
Call TakeScreenShot()
Call SelectTab("HEADER_DETAIL","Org. Data",False)
Wait(2)

'Enter the Organisation details
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call PressEnter() 
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)    
Call PressEnter()

'Enter Order Details'
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False)
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_Delivery_Date),False) 
Call PressEnter()
Wait(5)

Call ClickButtonIfExist("Continue",True)
wait(2)

Call ClickButtonIfExist("Continue",True)
wait(2)

'Capture the screenshot
Call TakeScreenShot()

Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_Delivery_Date),False)



'Click on Save Buton
Call SelectMenuBar("Purchase Order;Save")
Wait(2)

Call ClickButtonIfExist("Continue",True)
wait(2)

Call ClickButtonIfExist("Save",True)

'Validate If STO Retail Order is generated
Call GetStatusBar("item2","DT_STO_NUMBER_OUTPUT")
'Call VerifyStatusBar("STO Retail created under the number " & DT_STO_NUMBER_OUTPUT)

''----------------------Tcode ME23N----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_ME21N_0014_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ME21N_0014_OKCD)
Call TakeScreenShot()

Call ClickButtonIfExist("Other Purchase Order   \(Shift\+F5\)",False)
wait(2)
Call SetTextboxPopupIfExist("MEPO_SELECT-EBELN","Pur\. Order",DT_ME21N_0003_PUR_ORDER)
Call PressEnter()
Wait(2)
Call TakeScreenShot()

'Click on Messages button and get the output type.
Call ClickButtonIfExist("Messages   \(Shift\+F9\)",False)
wait(2)
Call TakeScreenShot()

'Verify the Output Type data
Call VerifyTableCellContent(1,"Output Type","SAPDV70ATC_NAST3",DT_ME21N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)


''''''''''''Lines of code from Line 136 to Line 170  implemented as part of the resolution provided for the defect 34372"""""""""""""""""""

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


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

