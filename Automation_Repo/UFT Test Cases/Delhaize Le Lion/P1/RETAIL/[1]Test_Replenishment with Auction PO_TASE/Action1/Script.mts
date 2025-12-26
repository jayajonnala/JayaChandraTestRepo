
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : [1]Test_Replenishment with Auction PO_TASE
'.................Test Scenario: AT_P2P_Fresh_Replenishment via Workbench and Auction PO
'.................TCode: ME21N
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

gstrTestCaseName = "Test_Replenishment with Auction PO_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode ME21N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()


Call SetCombo("MEPO_TOPLINE-BSART","Auction PO Retail")
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,FALSE)
Call SetTextbox("Purch\. Org\.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,FALSE)
Call SetTextbox("Purch\. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,FALSE)
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,FALSE)

' ClickButtonIfExist(tooltipOrButtonName, blnIsItPopup)
Call ClickButtonIfExist("Expand Items Ctrl\+F3",False)
Call ClickButtonIfExist("Switch Off Document Overview   \(F9\)",False)

'enter article table
Call SetTableData("SAPLMEGUITC_1211","Article",1, , ,DT_ME21N_1211_TABLECELL_ARTICLE_0,False)
Call SetTableData("SAPLMEGUITC_1211","PO Quantity",1, , ,DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date",1, , ,ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_0),False)
Call SetTableData("SAPLMEGUITC_1211","Plnt",1, , ,DT_ME21N_1211_TABLECELL_SITE_0,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

'validate latest GR date
Call SelectTab("ITEM_DETAIL", "Delivery", False)
Call TakeScreenShot()

' VerifyifGuiLabelExistsByRelativeid(Content, Relativeid)
Call VerifyifGuiLabelExistsByRelativeid(DT_ME21N_1313_CHECK_TEXT_OF_LATEST_GR_DATE,"wnd\[0\]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1313/lblMEPO1313-LEWED")
Call VerifyTextBoxContent("Latest GR Date","MEPO1313-LEWED", 0,ConvertDate(DT_ME21N_1313_CHECK_TEXT_OF_LATEST_GR_DATE_OCC1),False)

Call ClickButton("Save   \(Ctrl\+S\)", False)
Call TakeScreenShot()

'statusbar message
Call GetTextStatusBar("DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR_OUTPUT",DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

