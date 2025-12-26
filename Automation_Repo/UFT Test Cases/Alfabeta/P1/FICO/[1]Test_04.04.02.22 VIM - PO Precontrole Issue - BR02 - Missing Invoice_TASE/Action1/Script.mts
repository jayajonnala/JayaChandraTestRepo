

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04.04.02.22 VIM - PO Precontrole Issue - BR02 - Missing Invoice
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

gstrTestCaseName = "Test_04.04.02.22 VIM"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\P1_FICO\DT_04.04.02.22 VIM - PO Precontrole Issue - BR02 - Missing Invoice_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'DataRowSet =3
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

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

Call SetTextbox("Vendor Reference","/OPT/VIM_BL_1RIDX_OCR_DATA-XBLNR","",DT_OPTVIM_7AX2_2002_VENDOR_REFERENCE,False)
Call SetTextbox("Document Date","/OPT/VIM_BL_1RIDX_OCR_DATA-BLDAT","",FormatBlank(DT_OPTVIM_7AX2_2002_DOCUMENT_DATE),False)
Call SetTextbox("PO Number","/OPT/VIM_BL_1RIDX_OCR_DATA-EBELN","","",False)
Call TakeScreenShot()
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
Call ClickButtonToolBar("&MB_FILTER",0)
Call SelectCellGuiGrid("Column Set",0,5,"Column Name",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call SetTextbox("Document Id","%%DYN001-LOW","",DT_OPTVIM_7AX2_1105_DOCUMENT_ID,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(Enter\)",True)
 Wait 5
Call ClickButton("Refresh   \(F5\)",False)
Call TakeScreenShot()

 Wait 5

Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False)
Call ClickButton("Simulate and Bypass Business Rules   \(Ctrl\+Shift\+F9\)",False)
Call VerifyGridCellContent("",3,"Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_1_STATUS)
Call ClickButton("Exit   \(Shift\+F7\)",True)
Call SetTextbox("Document Date","GH_IDX_APPLICATION->MS_IDX_HEADER-BLDAT","",ConvertDate(DT_OPTVIM_7AX2_1100_DOCUMENT_DATE),False)
Call TakeScreenShot()
Call PressEnter()     ' 
   ' 

Call ClickButton("Simulate and Bypass Business Rules   \(Ctrl\+Shift\+F9\)",False)
Call VerifyGridCellContent("",3,"Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_1_STATUS_OCC1)
Call ClickButton("Exit   \(Shift\+F7\)",True)

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call LogOff()
Call FinalStatus ()



