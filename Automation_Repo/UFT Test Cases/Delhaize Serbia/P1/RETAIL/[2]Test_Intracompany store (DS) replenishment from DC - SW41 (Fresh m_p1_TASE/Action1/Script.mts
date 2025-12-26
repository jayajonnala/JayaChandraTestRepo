
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Intracompany store (DS) replenishment from DC - SW41 (Fresh m_p1
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
'.................Test Script Name : Test_Intracompany store (DS) replenishment from DC - SW41 (Fresh m_p1
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 11th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Intracompany store (DS) replenishment from DC - SW41 (Fresh m_p1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Intracompany store (DS) replenishment from DC - SW41 (Fresh m_p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode ME21N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
Call SetTextbox("Supplying Site","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_SUPPLYING_SITE,False) 
Call TakeScreenShot()
Call SetTextbox("Doc\. date","MEPO_TOPLINE-BEDAT","",ConvertDate(DT_ME21N_1105_DOC_DATE),False)

Call SelectTab("HEADER_DETAIL","Additional Data",False)
Call TakeScreenShot()
Call SelectTab("HEADER_DETAIL","Org. Data",False)
Wait(2)
Call TakeScreenShot()

Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call PressEnter() 
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)    
Call PressEnter()
Call TakeScreenShot()

Call ClickButtonIfExist("Expand Items Ctrl\+F3",False)
Call TakeScreenShot()

Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False)
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_ME21N_1105_DOC_DATE),False) 
Call TakeScreenShot()

Call SetTableData("SAPLMEGUITC_1211","Article","2","","",DT_ME21N_1211_TABLECELL_ARTICLE_1,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","2","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_1,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_ME21N_1211_TABLECELL_SITE_1,False)
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",ConvertDate(DT_ME21N_1105_DOC_DATE),False)
Call TakeScreenShot()

Call SetTableData("SAPLMEGUITC_1211","Article","3","","",DT_ME21N_1211_TABLECELL_ARTICLE_2,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","3","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_2,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","3","","",DT_ME21N_1211_TABLECELL_SITE_2,False)
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","3","","",ConvertDate(DT_ME21N_1105_DOC_DATE),False) 
Call TakeScreenShot()

Call SetTableData("SAPLMEGUITC_1211","Article","4","","",DT_ME21N_1211_TABLECELL_ARTICLE_3,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","4","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_3,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","4","","",DT_ME21N_1211_TABLECELL_SITE_3,False)
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","4","","",ConvertDate(DT_ME21N_1105_DOC_DATE),False)
Call TakeScreenShot()

Call PressEnter()
Call TakeScreenShot()
Call ClickButtonIfExist("Continue",True)
Call TakeScreenShot()
Call PressEnter()
Call PressEnter()
Call PressEnter()

Call GetTableCellData("SAPLMEGUITC_1211","Itm",1,"Article",DT_ME21N_1211_TABLECELL_ARTICLE_0,"DT_ME21N_1211_CHECK_TEXT_OF_TABLECELL_ITM_0",False)


Call PressEnter()

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call ClickButtonIfExist("SPOP-VAROPTION1",True)
Call TakeScreenShot()
Call GetStatusBar("item2","DT_STO_NUMBER_OUTPUT")
Call WriteRunTimeDataToExcel("DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1",DT_STO_NUMBER_OUTPUT)
Call VerifyStatusBar("STO Retail created under the number " &DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1)


''''''''''''Lines of code from Line 125 to Line 160  implemented as part of the resolution provided for the defect 34372"""""""""""""""""""
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

''------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

