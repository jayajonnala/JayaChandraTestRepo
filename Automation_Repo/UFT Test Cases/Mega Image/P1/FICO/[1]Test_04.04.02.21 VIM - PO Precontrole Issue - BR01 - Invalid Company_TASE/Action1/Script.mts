
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04.04.02.21 VIM - PO Precontrole Issue - BR01 - Invalid Company
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


gstrTestCaseName = "Test_04.04.02.21 BR01"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'----------------------Tcode /OPT/VIM_7AX2----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Arch\.","S_ARCHIV_ID","",DT_OPTVIM_7AX2_0002_ARCH,False)   
Call SetTextbox("Addl docs","S_DOCID","",DT_OPTVIM_7AX2_0002_ADDL_DOCS,False)   
Call TakeScreenShot()
Call ClickButton("Copy From",False)
Wait(2)

Call ClickButton("SUBMIT   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call GetStatusBar("item2","DT_DOCUMENT_ID_OUTPUT")
Call GetStatusBar("item4","DT_WORKFLOW_ID_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
wait (2)
Call VerifyStatusBar("Document Number "&DT_DOCUMENT_ID_OUTPUT&" Workflow ID "&DT_WORKFLOW_ID_OUTPUT)


'----------------------------Tcode /n/OPT/VIM_WP --------------------------------
Call SetTcode(DT_OPTVIM_7AX2_0002_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Document Id","H1_DOID-LOW","",DT_DOCUMENT_NO,False)
Call TakeScreenShot()
Call ClickButton("Apply",False)
Wait(2)
Call TakeScreenShot()

Call ClickButton("Switch Work View   \(Shift\+F6\)",False)
Call SelectRadioButtonIfPopupExists("SPOPLI-SELFLAG","All Users View")
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Wait(2)
'
Call TakeScreenShot()
Call ClickCellGuiGrid("All Inbox \(1  Hit\) - Selection Active",0,"Execute",1,"","",False)

Call SelectTab("TAB_MAIN","Basic Data",False)
Wait(1)
Call SetTextbox("Company Code","GH_IDX_APPLICATION->MS_IDX_HEADER-BUKRS",""," ",False)
Call TakeScreenShot()
Call PressEnter()   
Call PressEnter()   
Call PressEnter()   

Call ClickButton("Simulate and Bypass Business Rules   \(Ctrl\+Shift\+F9\)",False)
Wait(1)
Call TakeScreenShot()


'Call VerifyGridCellContentByRefColumn("","","Process Type","902","Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_0_STATUS)

Call VerifyGridCellContent("",DT_PROC_ROW_NUM,"Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_0_STATUS)

Call ClickButton("Exit   \(Shift\+F7\)",True)
Wait(2)

Call SetTextbox("Company Code","GH_IDX_APPLICATION->MS_IDX_HEADER-BUKRS","",DT_OPTVIM_7AX2_1100_COMPANY_CODE_OCC1,False)
Call TakeScreenShot()
Call PressEnter()   
Call PressEnter()   
Call PressEnter()   

Call ClickButton("Simulate and Bypass Business Rules   \(Ctrl\+Shift\+F9\)",False)
Wait(1)
Call TakeScreenShot()


'Call VerifyGridCellContentByRefColumn("","","Process Type","902","Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_0_STATUS_OCC1)


Call VerifyGridCellContent("",DT_PROC_ROW_NUM,"Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_0_STATUS_OCC1)

Call ClickButton("Exit   \(Shift\+F7\)",True)
Wait(2)

Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(2)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

