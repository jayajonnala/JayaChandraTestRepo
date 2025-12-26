
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "Test_06-09-68-Manual mvt 101check_SAP"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
gstrInputExcelFilePathAndName="S:\TASETestData\DLL\RETAIL\DT_06-09-68-Manual GR in WH (mvt type 101)_check_SAP_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''' Login '''
SAPGuiUtil.OpenConnection("R1E - SAP RETAIL Pre-Production EUROPE")
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''SAP Login'''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     ' - Line (12)
Call PressEnter()     ' - Line (13)

''INPUT''

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PUR_ORDER,True) ' - Line (9)
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True) ' - Line (10)
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4", false)
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)
'
' GetGridContent(gridTitle, gridIndex, columnName, rowNumber, refColumn, refFieldVal, dataTableColumnName)
'Call GetGridContent("","","Article Document",1,"Reference","",DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_output)
'Call GetGridContent("","","Article Document",2,"Reference","",DT_GET_ARTICLE_DOC_output)
' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
call VerifyGridCellContent("",2,"Article Document",0,DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
call VerifyGridCellContent("",2,"Movement type",0,DT_MVT_TYPE1)
'call VerifyGridCellContent("",2,"Movement type",0,DT_MVT_TYPE2)
'Call GetStatusBar("item1","DT_PO_NUMBER_OUTPUT")
Call TakeScreenShot
'
'
'
'
'
'
'
'
'
'
'Call SetComboByKey("MEPO_TOPLINE-BSART", DT_ME21N_1105_MEPO_TOPLINEBSART)
'Call ClickButtonIfExist("Expand Header Ctrl\+F2", False)
'Call SelectTab("HEADER_DETAIL","Org. Data",False)
'Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD","","EVG90198",False)         
'Call PressEnter()     
'Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False)     
'Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
'Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)        
'Call PressEnter()     
'Call SetTableData("SAPLMEGUITC_1211","Article","1","","","7154789",False)     
'Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)  
'Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_0),False) 
'Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False) 
'Call PressEnter()
'Call SetTableData("SAPLMEGUITC_1211","Article","2","","","7154792",False)     
'Call SetTableData("SAPLMEGUITC_1211","PO Quantity","2","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_1,False)     
'Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_1),False)     
'Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_ME21N_1211_TABLECELL_SITE_1,False)     
'Call TakeScreenShot
'Call PressEnter()
'Call TakeScreenShot
'Call ClickButton("Save   \(Ctrl\+S\)",False)
'Call ClickButtonIfExist("Save",True)
'Call TakeScreenShot
'Call GetStatusBar("item2","DT_PO_NUMBER_OUTPUT")
'Call VerifyStatusBar("Standard PO Retail created under the number "&DT_PO_NUMBER_OUTPUT)
'
'
'
'
'
'
'''''''''''''''''''''''
'
Call LogOff()
Call FinalStatus ()
'''control data

