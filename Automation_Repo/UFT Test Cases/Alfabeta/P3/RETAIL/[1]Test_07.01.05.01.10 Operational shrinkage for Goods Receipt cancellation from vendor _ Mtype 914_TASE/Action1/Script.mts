

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_07.01.05.01.10 Operational shrinkage for Goods Receipt cancellation from vendor _ Mtype 914
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


gstrTestCaseName = "Test_07.01.05.01.10 Operational"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''--------------------------------MEK3-----------------------------
Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_MEK3_0100_CONDITION_TYPE,False)
Call TakeScreenShot()
Call PressEnter()     ' 
Call TakeScreenShot()
Call SetTextbox("Purch\. Organization","F001","",DT_MEK3_1000_PURCH_ORGANIZATION,False)
Call SetTextbox("Article","F002-LOW","",DT_MEK3_1000_ARTICLE,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call GetTableCellData("SAPMV13ATCTRL_FAST_ENTRY","Amount",1,"","","DT_MEK3_1993_GET_TEXT_OF_TABLECELL_AMOUNT_0_OUTPUT",false)
Call ClickButton("Exit   \(Shift\+F3\)",False)


''--------------------------------mb51-----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MEK3_0100_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_MEK3_0100_OKCD)

Call SetTextbox("Article","MATNR-LOW","",DT_MEK3_1000_ARTICLE_OCC1,False)
Call SetTextbox("Site","WERKS-LOW","",DT_MEK3_1000_SITE,False)
Call SetTextbox("Movement type","BWART-LOW","",DT_MEK3_1000_MOVEMENT_TYPE,False)
Call SetTextbox("Company Code","BUKRS-LOW","",DT_MEK3_1000_COMPANY_CODE,False)

'Call SetTextbox("Reference","XBLNR-LOW","",FormatBlank(DT_MB51_REFERENCE_INPUT),False)
'Call SetTextboxNoLabel("XBLNR-LOW",0,FormatBlank(DT_MB51_REFERENCE_INPUT),False)

Call SetTextboxNoLabel("BKTXT-LOW",0,FormatBlank(DT_MEK3_1000_DOCUMENT_HEADER_TEXT),False)
Call SetTextboxNoLabel("MBLNR-LOW",0,FormatBlank(DT_MEK3_1000_ARTICLE_DOCUMENT),False)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False) 
Call TakeScreenShot()
Call ClickButton("Detail List   \(Ctrl\+Shift\+F12\)",False) 
Call VerifyGridCellContent("",1,"ERFMG",0,DT_MEK3_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ERFMG)
Call GetGridContentByTitle("",0,"MBLNR",1,"DT_MEK3_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MBLNR_OUTPUT")

Call LogOff()
Call FinalStatus ()


