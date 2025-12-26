
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Purchasing - Storno in reversal order SW31 dry goods_p1
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
'.................Test Script Name : Test_Purchasing - Storno in reversal order SW31 dry goods_p1
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 19th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Purchasing - Storno in reversal order SW31 dry goods_p1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Purchasing  - Storno in reversal order SW31 dry goods_p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode MR8M----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextboxNoLabel("RBKPV-BELNR","",DT_DOC_NO,False)
Call SetTextbox("Fiscal Year","RBKPV-GJAHR","",Year(Date),False)   
Call SetTextbox("Reversal Reason","UF05A-STGRD","",DT_OPTVIM_WP_0300_REVERSAL_REASON,False) 
Call TakeScreenShot()

Call ClickButton("Reverse   \(Ctrl\+S\)",False)
Wait(2)
Call TakeScreenShot()

Call GetStatusBar("item1","DT_REVERSAL_DOC_NO_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar("Document reversed with no. "&DT_REVERSAL_DOC_NO_OUTPUT&": Please manually clear FI documents")
Call SetTextbox("Invoice Document No\.","RBKPV-BELNR","",DT_OPTVIM_WP_0300_CHECK_TEXT_OF_STATUSBAR,False)
Call TakeScreenShot()
Call ClickButton("Display Document   \(F2\)",False)
Wait(2)
Call TakeScreenShot()
Call ClickButtonIfExist("Follow-On Documents \.\.\.   \(F8\)",False)
wait(2)

Call TakeScreenShot()

''----------------------Tcode MIGO----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_OPTVIM_WP_0750_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_OPTVIM_WP_0750_OKCD)

Call SetCombo("GODYNPRO-ACTION","Cancellation")
Call SetComboByKey("GODYNPRO-REFDOC","R02")
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_GODYNPROMAT_DOC,False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,Year(Date),False)
Call TakeScreenShot()
Call PressEnter() 
Wait(2)
Call TakeScreenShot()
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call GetStatusBar("item1","DT_CANCELLATION_DOC_NO_OUTPUT")
Call VerifyStatusBar("Article document " & DT_CANCELLATION_DOC_NO_OUTPUT&" "&"posted")

Call ClickButton("Exit   \(Shift\+F3\)",False)
Wait(1)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
