

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04.04.02.04 VIM - NPO Precontrole Issue - BR04 - Invalid Vendor (NPO)
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


gstrTestCaseName = "Test_04.04.02.04-BR04"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\P1_FICO\DT_04.04.02.04 VIM - NPO Precontrole Issue - BR04 - Invalid Vendor_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =4
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()
'
'--------------------------------------------  /n/OPT/VIM_7AX2----------------------------------------------



Call SetTextbox("Arch\.","S_ARCHIV_ID","",DT_OPTVIM_7AX2_0002_ARCH,False)
Call SetTextbox("Addl docs","S_DOCID","",DT_OPTVIM_7AX2_0002_ADDL_DOCS,False)
Call TakeScreenShot()
Call ClickButton("Copy From",False)
Call TakeScreenShot()

Call SetTextbox("Vendor Reference","/OPT/VIM_BL_1RIDX_OCR_DATA-XBLNR","",FormatBlank(DT_OPTVIM_7AX2_2002_VENDOR_REFERENCE),False)
Call SetTextbox("PO Number","/OPT/VIM_BL_1RIDX_OCR_DATA-EBELN","",FormatBlank(DT_OPTVIM_7AX2_2002_PO_NUMBER),False)
Call SetTextbox("Vendor Number","/OPT/VIM_BL_1RIDX_OCR_DATA-LIFNR","",FormatBlank(DT_OPTVIM_7AX2_2002_VENDOR_NUMBER),False)
Call TakeScreenShot()
Call SelectRowGuiTableByRow("/OPT/VIM_BL_M1_SUBSCREENSTC_INDEX_ITEMS",1,False)
Call ClickButton("Delete line",False)
Call ClickButton("SUBMIT   \(F8\)",False)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Call GetStatusBar("item2","DT_OPTVIM_7AX2_0002_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item4","DT_OPTVIM_7AX2_0002_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_OPTVIM_7AX2_0002_CHECK_TEXT_OF_STATUSBAR)

'--------------------------------------------  /n/OPT/VIM_7AX2----------------------------------------------


Call SetTcode(DT_OPTVIM_7AX2_0002_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Document Id","H1_DOID-LOW","",DT_OPTVIM_7AX2_0841_SEARCH_TERM,False)
Call ClickButton("BT_H_APPLY",False)
Call ClickButton("Switch Work View   \(Shift\+F6\)",False)
Call TakeScreenShot()
Call SelectRadioButton("SPOPLI-SELFLAG","All Users View",True)
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Hide Detail Pane   \(Ctrl\+F2\)",False)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot()
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_OPTVIM_7AX2_0841_SEARCH_TERM_OCC1,True)
Call TakeScreenShot()
Call ClickButton("OK   \(Enter\)",True)
Call TakeScreenShot()
Call ClickButton("Cancel   \(F12\)",True)

Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Show Options   \(Ctrl\+F12\)",False)
Call ClickCellGuiGrid("",0,"#1",1,"","",False)
Call SetComboByKey("G_NEW_DOC_TYPE",DT_OPTVIM_7AX2_0020_NEW_DOCUMENT_TYPE)
Call TakeScreenShot()
Call ClickButton("Continue   \(F5\)",True)
 Wait 5

Call ClickButton("Refresh   \(F5\)",False)
 Wait 5

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot()
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_OPTVIM_7AX2_0841_SEARCH_TERM_OCC1,True)
Call TakeScreenShot()
Call ClickButton("OK   \(Enter\)",True)
Call TakeScreenShot()
Call ClickButton("Cancel   \(F12\)",True)

Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False)
Call ClickButton("Simulate and Bypass Business Rules   \(Ctrl\+Shift\+F9\)",False)
Call VerifyGridCellContentByRefColumn("","","PROCESS_TYPE","201","Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_3_STATUS)
'Call VerifyGridCellContent("",4,"Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_3_STATUS)
Call ClickButton("Exit   \(Shift\+F7\)",True)
Call SetTextbox("Vendor Number","GH_IDX_APPLICATION->MS_IDX_HEADER-LIFNR","",DT_OPTVIM_7AX2_9100_VENDOR_NUMBER,False)
Call TakeScreenShot()
Call PressEnter()   ' 

Call ClickButton("Simulate and Bypass Business Rules   \(Ctrl\+Shift\+F9\)",False)
'Call VerifyGridCellContent("",4,"Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_3_STATUS_OCC1)
Call VerifyGridCellContentByRefColumn("","","PROCESS_TYPE","201","Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_3_STATUS_OCC1)
Call ClickButton("Exit   \(Shift\+F7\)",True)
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBar(DT_OPTVIM_7AX2_1000_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus ()


