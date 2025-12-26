
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.1.1.14.1.-Maintain Substitution of Article
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

gstrTestCaseName = "Test_2.1.1.14.1.-Maintain Substitution of Article"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.7.1.2.2. Create Low Level Header ZSSP - Sales Support Promo.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=4
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 

''''---------Login----------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
Call TakeScreenShot()
'----------------------Tcode WRFFUART----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()

Call ClickButton("Article Selection   \(F6\)",False)
Call SetTextbox("Substit\. Profile","P_FUTYP1","",DT_WRFFUART_0103_SUBSTIT_PROFILE,True)
Call SetTextbox("Article","S_ARTNR-LOW","",DT_WRFFUART_0103_ARTICLE,True)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)
Call TakeScreenShot()
Call SelectRowGuiGridbyRowNo("Article Selection", 0, 1, False)
Call ClickButtonToolBar("ASSIGN", 0)
Call TakeScreenShot()

Call SetGridData("Substitution Type: 10 DS Replacement", 1, "ORG_ART_FACTOR", DT_WRFFUART_0100_GRIDCELL_0_FACTOR, False)
Call SetGridData("Substitution Type: 10 DS Replacement", 1, "SUBST_ART_FACTOR", DT_WRFFUART_0100_GRIDCELL_0_FACTOR_OCC1, False)
Call SetGridData("Substitution Type: 10 DS Replacement", 1, "FOLUP_ART_NR", DT_WRFFUART_0100_GRIDCELL_0_FUREPLACEMENT_ART, False)
Call SetGridData("Substitution Type: 10 DS Replacement", 1, "ASORT", DT_WRFFUART_0100_GRIDCELL_0_ASSORTMENT, False)
Call SetGridData("Substitution Type: 10 DS Replacement", 1, "PRIORITY_A", DT_WRFFUART_0100_GRIDCELL_0_PRIORITY, False)

Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Wait 5
Call VerifyTextBoxContent("Information Message","MESSTXT1", "", LCase(DT_WRFFUART_0010_CHECK_TEXT_OF_MESSTXT1), True)
Call ClickButton("Continue   \(Enter\)",True)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

