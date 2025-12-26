
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Returns from DC - CC RS01_p1_TASE
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


gstrTestCaseName = "Test_Returns from DC - CC RS01_p1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName=""
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode MIRO----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()   
Call TakeScreenShot
Call SetComboByKey("MEPO_TOPLINE-BSART", DT_ME21N_1105_MEPO_TOPLINEBSART)
Call ClickButtonIfExist("Expand Header Ctrl\+F2", False)
Call SelectTab("HEADER_DETAIL","Org. Data",False)
'Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Call PressEnter()     
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False)     
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)        
Call PressEnter() 
Call TakeScreenShot
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False)     
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)  
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","Returns Item","1","","","ON",False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
'''''''Call SetTableData("SAPLV69ATCTRL_KONDITIONEN","Amount","1","Name","Gross Price","100",False)  
'''''''Call PressEnter()


Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
Call TakeScreenShot
Call GetStatusBar("item2","DT_ME21N_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Standard PO Retail created under the number "&DT_ME21N_STATUSBAR_OUTPUT)

Call SetTcode(DT_ME21N_0014_OKCD) 
Call PressEnter()   
Call TakeScreenShot
Call ClickButton("Other Purchase Order   \(Shift\+F5\)",false)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME21N_STATUSBAR_OUTPUT,True) ' - Line (9)
'''Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True) ' - Line (10)
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Messages   \(Shift\+F9\)",False)
Call TakeScreenShot


Call LogOff()
Call FinalStatus ()




'*********************************************End Of Script*********************************************************************

