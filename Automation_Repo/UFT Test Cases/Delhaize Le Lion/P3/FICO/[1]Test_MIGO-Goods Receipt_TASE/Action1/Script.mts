
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :[1]Test_MIGO-Goods Receipt
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_MIGO-Goods Receipt_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\DLL\TASE_DT_MIGO-Goods Receipt.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//



Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     ' - Line (2)
Call PressEnter()     ' - Line (3)

Call SetCombo("GODYNPRO-ACTION","Goods Receipt")
Call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetTextBoxNoLabel("GODYNPRO-PO_NUMBER","0",DT_PO,False)


Call PressEnter()
 
 Call TakeScreenShot()
 
 Call SetTextBox("Delivery Note","GOHEAD-LFSNR",0,DT_MIGO_0110_DELIVERY_NOTE,False)
 Call PressEnter()
 
'Call SetTextboxByGuiLabelXYCordDiff("Document Date","0",-350,0,DT_MIGO_0110_DELIVERY_NOTE)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE", 0, "ON", False)

Call ClickButton("Check Entries   \(F7\)",False)

Call ClickButton("btn\[11\]",False)     ' - Line (14)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_ART_DOC_OUTPUT")

Call VerifyStatusBar("Article document "&DT_ART_DOC_OUTPUT&" posted")

Call LogOff()

Call FinalStatus ()

