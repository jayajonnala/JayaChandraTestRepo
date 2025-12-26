

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02LIV12_012_Credit_Note_DSD_w_price_change_necessary
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
    RunTimeResultFolder= Parameter("RunTimeResultFolder")    
End If

gstrTestCaseName = "Test_02LIV12_012_Credit_Note_DSD_w_price_change_necessary"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'''
'''
'''--------------------------------MR21-----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_NUM",Cint(DT_NUM)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","MR21HEAD-BUKRS","",DT_MR21_201_COMPANY_CODE,False)
Call SetTextbox("Site","MR21HEAD-WERKS","",DT_MR21_201_SITE,False)
Call SetTextbox("Reference","MR21HEAD-XBLNR","",DT_MR21_201_REFERENCE,False)
Call TakeScreenShot()
Call PressEnter()   

Call TakeScreenShot()
Call SetTableData("SAPRCKM_MR21MR21_TABCONTROL","Article","1","","",DT_MR21_0250_TABLECELL_ARTICLE_0,False) 
Call PressEnter()   
Call  GetTableCellData("SAPRCKM_MR21MR21_TABCONTROL","Current valuation price",1,"","","DT_GET_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTableData("SAPRCKM_MR21MR21_TABCONTROL","New price","1","","",DT_MR21_0250_TABLECELL_NEW_PRICE_0,False) 
Call PressEnter()   
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S")

Call GetStatusBar("item1","DT_DOC_NUM_OUTPUT")
Call VerifyStatusBar("Price change document " & DT_DOC_NUM_OUTPUT &" posted")

Call ClickButton("Display Document   \(F7\)",False)
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call PressEnter()   

Call SelectRowGuiGridbyRowNo("",0,1,False)
Call ClickButton("Accounting Documents\.\.\.   \(F6\)",False)

Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()


