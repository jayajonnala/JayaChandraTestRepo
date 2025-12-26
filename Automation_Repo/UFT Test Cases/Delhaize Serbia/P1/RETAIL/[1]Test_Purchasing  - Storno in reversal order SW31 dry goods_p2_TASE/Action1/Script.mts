
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Purchasing - Storno in reversal order SW31 dry goods_p2
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
'.................Test Script Name : Test_Purchasing - Storno in reversal order SW31 dry goods_p2
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 19th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Purchasing - Storno in reversal order SW31 dry goods_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Purchasing  - Storno in reversal order SW31 dry goods_p2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode MIGO----------------------------

'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
'
Call SetCombo("GODYNPRO-ACTION","Display")
Call SetComboByKey("GODYNPRO-REFDOC","R02")

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_MIGO_2010_GODYNPROMAT_DOC,False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,Year(Date),False)
Call TakeScreenShot()
Call PressEnter() 
Wait(2)
Call TakeScreenShot()

'Navigate to the Document Details
Call SelectTab("TS_GOHEAD","Doc. info",False)
Wait(1)
Call TakeScreenShot()
Call ClickButton("FI Documents",False)
Wait(5)
Call TakeScreenShot()

'Verify the Title
Call CheckScreen(DT_SAPTRANSACTIONCODE,"Display Document: Data Entry View")

''----------------------Tcode ZMDPU_COPY_DELETE_PO----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MIGO_1000_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MIGO_1000_OKCD)

'Enter the PO 
Call SetTextbox("Purchasing Document","P_EBELN","",DT_MIGO_1000_PURCHASING_DOCUMENT,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()

'Verify GUI Label
Call VerifyifGuiLabelExists(DT_MIGO_0120_CHECK_TEXT_OF_RELEASE)


'Click on Continue
Call ClickButton("Continue   \(Enter\)",False) 
Wait(2)

'Get the Document No
Call GetTextboxValue("MEPO_TOPLINE-EBELN",0,"DT_DOCUMENT_NUMBER",False)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

