
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.6.1.7.1. Check Price Change File from Workflow_V2
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

gstrTestCaseName = "Test_2.6.1.7.1. Check Price Change File from Workflow_V2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.6.1.7.1. Check Price Change File from Workflow.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode VKP5 ----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()

Call ClickButton("Get variant\.\.\.   \(Shift\+F5\)",False)
Call TakeScreenShot()

Call SetTextbox("Variant","V-LOW","",DT_VKP5_0100_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButton("Execute   \(F8\)",True)
Call TakeScreenShot()


Call SetTextbox("Article","S_MATNR-LOW","",DT_VKP5_1000_ARTICLE,False)
Call SetTextbox("Validity","S_DATUM-LOW","",ConvertDate(DT_VKP5_1000_VALIDITY),False)
Call SetTextbox("to","S_DATUM-HIGH","",ConvertDate(DT_VKP5_1000_TO),False)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call SetGridData("", 1, "SPANE", DT_VKP5_0100_GRIDCELL_0_MARG_NET,False)
Call SetGridData("", 2, "SPANE", DT_VKP5_0100_GRIDCELL_1_MARG_NET,False)
Call SetGridData("", 3, "SPANE", DT_VKP5_0100_GRIDCELL_2_MARG_NET,False)
Call SetGridData("", 4, "SPANE", DT_VKP5_0100_GRIDCELL_3_MARG_NET,False)
Call SetGridData("", 5, "ENDPR", DT_VKP5_0100_GRIDCELL_4_FINAL_PRICE,False)
Call SetGridData("", 6, "ENDPR", DT_VKP5_0100_GRIDCELL_5_FINAL_PRICE,False)
Call SetGridData("", 7, "ENDPR", DT_VKP5_0100_GRIDCELL_6_FINAL_PRICE,False)
Call TakeScreenShot()

Call SelectRowGuiGridbyRowNo("", 0, 1, False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()


Call SelectRowGuiGridbyRowNo("", 0, 1, False)
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait 5

Call ClickButtonifExist("Enter   \(F5\)",True)
Wait 5

Call GetStatusBar("item1","DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Data saved; pricing document "&DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" created")
Call WriteRunTimeDataToExcelGlobalSheet("DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call LogOff'
Call FinalStatus ()
