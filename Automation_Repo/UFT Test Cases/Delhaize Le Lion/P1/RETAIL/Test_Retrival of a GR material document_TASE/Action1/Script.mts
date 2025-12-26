
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Retrival of a GR material Document Number_TASE
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



gstrTestCaseName = "TC_05_Test_Retrival of a GR material Document Number_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_ABI067_001 Create ZCXT Retail Customer Local or Foreign_TASE.xls"

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''----------------------Tcode MIGO----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()


Call SetTextboxNoLabel("MBLNR-LOW", 0, DT_MATERIAL, False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call Clickbuttonifexist("Detail List   \(Ctrl\+Shift\+F12\)",False)
Call TakeScreenShot()
Call VerifyGridCellContentbyName("shell", 1, "Purchase order", "", DT_PO_NUMBER)
Call VerifyGridCellContentbyName("shell", 1, "Site", "", DT_SITE)
Call VerifyGridCellContentbyName("shell", 1, "Article", "", DT_ARTICLE)
Call VerifyGridCellContentbyName("shell", 1, "Material Description", "", DT_MATERIAL_DESCRIPTION)

Call  ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

Call  ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus()
