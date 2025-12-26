

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04.04.02.05 VIM - NPO Precontrole Issue - BR05a - Suspected Dupl
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



gstrTestCaseName = "Test_04.04.02.05 VIM-NPO Precontrole Issue-BR05a "
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\P1_FICO\DT_04.04.02.05 VIM - NPO Precontrole Issue - BR05a - Suspected Dupl_TASE.xls"


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

Call SetTextbox("Arch\.","S_ARCHIV_ID","",DT_OPTVIM_7AX2_0002_ARCH,False)
Call SetTextbox("Posting Date","/OPT/VIM_BL_1RIDX_OCR_DATA-BUDAT","",FormatBlank(DT_OPTVIM_7AX2_2002_POSTING_DATE),False)
Call SetTextbox("PO Number","/OPT/VIM_BL_1RIDX_OCR_DATA-EBELN","",FormatBlank(DT_OPTVIM_7AX2_2002_PO_NUMBER),False)
Call TakeScreenShot()
Call SelectRowGuiTableByRow("/OPT/VIM_BL_M1_SUBSCREENSTC_INDEX_ITEMS",1,False)
Call ClickButton("Delete line",False)
Call ClickButton("SUBMIT   \(F8\)",False)
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True)
Call GetTextStatusBar("DT_OPTVIM_7AX2_0002_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_DOCUMENT_CHECK)



'--------------------------------------------  /n/OPT/VIM_7AX2----------------------------------------------


Call SetTcode(DT_OPTVIM_7AX2_0002_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Document Id","H1_DOID-LOW","",DT_OPTVIM_7AX2_1105_DOCUMENT_ID,False)
Call ClickButton("BT_H_APPLY",False)
Call ClickButton("Switch Work View   \(Shift\+F6\)",False)
Call TakeScreenShot()
Call SelectRadioButton("SPOPLI-SELFLAG","All Users View",True)
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Hide Detail Pane   \(Ctrl\+F2\)",False)
Call ClickButtonToolBar("&MB_FILTER",0)
Call SelectCellGuiGrid("Column Set",0,5,"Column Name",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call SetTextbox("Document Id","%%DYN001-LOW","",DT_OPTVIM_7AX2_1105_DOCUMENT_ID,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(Enter\)",True)
 Wait 5
Call ClickButton("Refresh   \(F5\)",False)
 Wait 5

Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False)
Call ClickButton("Simulate and Bypass Business Rules   \(Ctrl\+Shift\+F9\)",False)
Call VerifyGridCellContent("",5,"Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_4_STATUS)
Call VerifyGridCellContent("",5,"PROCESS_TYPE",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_4_PROCESS_TYPE)
Call ClickButton("Exit   \(Shift\+F7\)",True)

Call ClickButton("Document is Duplicate   \(Shift\+F9\)",False)
Call TakeScreenShot()
Call ClickButton("Yes",True)
Call TakeScreenShot()
Call SetTextArea(DT_OPTVIM_7AX2_0011_TEXTEDIT_SHELL)
Call ClickButton("Save   \(Ctrl\+S\)",True)
Call ClickButton("Refresh   \(F5\)",False)
Call TakeScreenShot()

'--------------------------------------------  /n/opt/vim_va2----------------------------------------------


Call SetTcode(DT_OPTVIM_7AX2_1000_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Document Processing Number","S_DOCID-LOW","",DT_OPTVIM_7AX2_1000_DOCUMENT_PROCESSING_NUMBER,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call VerifyGridCellContent("Results \(1 Hit\)",1,"OVERALL_STATUS_TEXT",0,DT_OPTVIM_7AX2_2000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_OVERALL_STATUS_TEXT)


Call LogOff()
Call FinalStatus ()



