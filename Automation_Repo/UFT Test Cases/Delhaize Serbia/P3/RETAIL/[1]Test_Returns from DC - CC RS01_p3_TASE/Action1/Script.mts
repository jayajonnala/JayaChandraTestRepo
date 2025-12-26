
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Returns from DC - CC RS01_p3
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


gstrTestCaseName = "Test_Returns from DC - CC RS01_p3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DS\RETAIL\DT_Returns from DC - CC RS01_p3_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_DELIVERY_NOTE",DT_INCREMENT_DELIVERY_NOTE+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''----------------------Tcode MIRO----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()  

Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION)
Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
Call PressEnter()
Call TakeScreenShot 
Call SetTextBoxNoLabel("GODYNPRO-PO_NUMBER","0",DT_MIGO_2000_GODYNPROPO_NUMBER,False)
Call TakeScreenShot
Call PressEnter() 
Call SetTextbox("Delivery Note","GOHEAD-LFSNR",1,DT_MIGO_0110_DELIVERY_NOTE,False)
Call PressEnter() 
Call TakeScreenShot
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE", 0, "ON", False)
Call ClickButton("Post Document   \(Shift\+F11\)",false)
Call GetStatusBar("item1","DT_MIGO_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Article document "&DT_MIGO_STATUSBAR_OUTPUT&" posted")
Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION_OCC1)
Call SetTextBoxNoLabel("GODYNPRO-MAT_DOC","0",DT_MIGO_STATUSBAR_OUTPUT,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call SelectTab("TS_GOHEAD","Doc. info", False)
Call ClickButtonIfExist("FI Documents",False)
Call TakeScreenShot


Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

