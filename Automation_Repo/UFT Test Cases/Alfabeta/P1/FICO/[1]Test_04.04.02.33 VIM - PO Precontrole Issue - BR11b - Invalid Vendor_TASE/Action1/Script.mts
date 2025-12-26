'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :[1]Test_04.04.02.33 VIM - PO Precontrole Issue - BR11b - Invalid Vendor_TASE
'.................Author : TCS_Ramesh
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



gstrTestCaseName = "Test_04.04.02.33 VIM"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\rtod\Documents\Input Datasheet\DLL\DT_04.04.02.12 VIM - NPO Precontrole Issue - BR10b - Invalid Vendor_TASE2.xls"

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'DataRowSet =2

Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
Wait(2)

'----------------------Tcode /OPT/VIM_7AX2----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call TakeScreenShot()
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()
 
Call SetTextbox("Arch\.","S_ARCHIV_ID","",DT_OPTVIM_7AX2_0002_ARCH,False)
Call SetTextbox("Addl docs","S_DOCID","",DT_OPTVIM_7AX2_0002_ADDL_DOCS,False)
Call TakeScreenShot()

Call ClickButton("Copy From",fALSE)
Call SetTextbox("Vendor Reference","/OPT/VIM_BL_1RIDX_OCR_DATA-XBLNR","",DT_OPTVIM_7AX2_2002_VENDOR_REFERENCE,False)
Call GetTextboxValue("/OPT/VIM_BL_1RIDX_OCR_DATA-VENDOR_TAX_NO","","DT_GET_VENDOR_TAX_Output",False)


Call SetTextbox("Arch\.","S_ARCHIV_ID","",DT_OPTVIM_7AX2_0002_ARCH,False)
Call SetTextbox("Vendor Reference","/OPT/VIM_BL_1RIDX_OCR_DATA-XBLNR","",DT_OPTVIM_7AX2_2002_VENDOR_REFERENCE_OCC1,False)
Call SetTextbox("Vendor VAT No","/OPT/VIM_BL_1RIDX_OCR_DATA-VENDOR_VAT_NO","",DT_OPTVIM_7AX2_2002_VENDOR_VAT_NO,False)
Call SetTextbox("Vendor Tax No","/OPT/VIM_BL_1RIDX_OCR_DATA-VENDOR_TAX_NO","","",False)
Call SetTextbox("Recepient VAT No","/OPT/VIM_BL_1RIDX_OCR_DATA-RECEPIENT_VAT_NO","",DT_OPTVIM_7AX2_2002_RECEPIENT_VAT_NO,False)
Call SetTextbox("PO Number","/OPT/VIM_BL_1RIDX_OCR_DATA-EBELN","",DT_OPTVIM_7AX2_2002_PO_NUMBER,False)


'Click on Submit Button
Call ClickButton("SUBMIT   \(F8\)",False)
Wait(2)

'Click on Continue
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(2)

'Validate If Purchase order is generated
Call GetStatusBar("item2","DT_DOCUMENT_ID_OUTPUT")
Call GetStatusBar("item4","DT_WORKFLOW_ID_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar("Document Number "&DT_DOCUMENT_ID_OUTPUT&" Workflow ID "&DT_WORKFLOW_ID_OUTPUT)





'''------------------------------TCode /n/OPT/VIM_WP-------------------------------------------------
Call SetTcode(DT_OPTVIM_7AX2_0002_OKCD)
Call PressEnter()


Call SetTextbox("Document Id","H1_DOID-LOW","",DT_OPTVIM_7AX2_1105_DOCUMENT_ID,False)
Call ClickButton("BT_H_APPLY",False)
Call ClickButton("Switch Work View   \(Shift\+F6\)",False)
Call SelectRadioButton("SPOPLI-SELFLAG","All Users View",True)
Call ClickButton("Continue   \(Enter\)",True)
Wait(2)
Call TakeScreenShot()


Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False)
Call TakeScreenShot()

'Selecting Process Tab
Call SelectTab("TAB_MAIN","Process",False)
Call TakeScreenShot()

'Selecting Tax Tab
Call SelectTab("TAB_MAIN","Tax",False)
Call TakeScreenShot()

'Click on Simulate and Bypass Business Rules
Call ClickButton("Simulate and Bypass Business Rules   \(Ctrl\+Shift\+F9\)",False)
Wait(1)
Call TakeScreenShot()


'Verify the Grid data
Call VerifyGridCellContentByRefColumn("","","PROCESS_TYPE","922","Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_9_STATUS)


'Click on Exit
Call ClickButton("Exit   \(Shift\+F7\)",False)
Wait(1)

'Setting Vendor Tax No Value
Call SetTextbox("Vendor Tax Number","GH_IDX_APPLICATION->MS_IDX_HEADER-VENDOR_TAX_NO","",DT_GET_VENDOR_TAX,False)
Call TakeScreenShot()
Call PressEnter()
Wait(1)
Call TakeScreenShot()

'Click on Simulate and Bypass Business Rules
Call ClickButton("Simulate and Bypass Business Rules   \(Ctrl\+Shift\+F9\)",False)
Wait(1)

'Verify the Grid data
Call VerifyGridCellContentByRefColumn("","","PROCESS_TYPE","922","Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_9_STATUS_OCC1)
Call TakeScreenShot()

'Click on Exit
Call ClickButton("Exit   \(Shift\+F7\)",False)
Wait(1)
Call TakeScreenShot()

'Save the data
Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(1)
Call VerifyStatusBar(DT_OPTVIM_7AX2_1000_CHECK_TEXT_OF_STATUSBAR)


'Click on Exit
Call ClickButton("Exit   \(Shift\+F3\)",False) 
Wait(1)
'Click on Exit
Call ClickButton("Exit   \(Shift\+F3\)",False) 
Wait(1)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()



'
'
'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [13,015294]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



