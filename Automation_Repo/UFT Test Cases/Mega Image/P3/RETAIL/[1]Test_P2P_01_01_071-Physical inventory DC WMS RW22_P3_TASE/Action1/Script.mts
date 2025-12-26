'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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


'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_P2P_01_01_071-Physical inventory DC WMS RW22_P3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 8th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_071-WMS RW22_P3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_P2P_01_01_071-Physical inventory DC WMS RW22_P3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''Login'''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
''''----------------------Tcode MI24----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Article","IM_MATNR-LOW","",DT_MI24_1000_ARTICLE,False)
Call SetTextbox("Site","IM_WERKS-LOW","",DT_MI24_1000_SITE,False)
SAPGuiSession("transaction:=MI24").SAPGuiWindow("transaction:=MI24").SAPGuiButton("tooltip:=Multiple selection","index:=4").Click
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_MI24_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_MI24_3010_TABLECELL_SINGLE_VALUE_1,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
Wait(1)

'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)

'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckBoxByGuiLabel(DT_MI24_3010_TABLECELL_SINGLE_VALUE_0,"17","1",DT_MI24_0120_NO_NAME)
Call ClickLabel(DT_MI24_3010_TABLECELL_SINGLE_VALUE_0,"",False)

'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Physical Inventory History   \(F2\)",False)

'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Article","ISEG-MATNR","",DT_MI24_0710_CHECK_TEXT_OF_ARTICLE,False)
Call VerifyTextBoxContent("Site","IKPF-WERKS","",DT_MI24_0710_CHECK_TEXT_OF_SITE,False)
Call VerifyTextBoxContent("Storage Loc\.","IKPF-LGORT","",DT_MI24_0710_CHECK_TEXT_OF_STOR_LOCATION,False)
Call VerifyTextBoxContent("Phys\.inv\.status","T064T-STEXT","",Lcase(DT_MI24_0710_CHECK_TEXT_OF_PHYSINVSTATUS),False)
Call VerifyTextBoxContent("Difference qty","VM07I-DIFMG","",DT_MI24_0710_CHECK_TEXT_OF_DIFFERENCE_QTY,False)

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)

Call SelectCheckBoxByGuiLabel(DT_MI24_3010_TABLECELL_SINGLE_VALUE_0,"17","1",DT_MI24_0120_NO_NAME_OCC1)
Call SelectCheckBoxByGuiLabel(DT_MI24_3010_TABLECELL_SINGLE_VALUE_1,"17","1",DT_MI24_0120_NO_NAME_OCC2)
Call ClickLabel(DT_MI24_3010_TABLECELL_SINGLE_VALUE_1,"",False)

'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Physical Inventory History   \(F2\)",False)

'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Site","IKPF-WERKS","",DT_MI24_0710_CHECK_TEXT_OF_SITE_OCC1,False)
Call VerifyTextBoxContent("Storage Loc\.","IKPF-LGORT","",DT_MI24_0710_CHECK_TEXT_OF_STOR_LOCATION_OCC1,False)
Call VerifyTextBoxContent("Phys\.inv\.status","T064T-STEXT","",Lcase(DT_MI24_0710_CHECK_TEXT_OF_PHYSINVSTATUS_OCC1),False)
Call VerifyTextBoxContent("Difference qty","VM07I-DIFMG","",DT_MI24_0710_CHECK_TEXT_OF_DIFFERENCE_QTY_OCC1,False)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

