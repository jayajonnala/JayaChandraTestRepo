

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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_02GR04_002_Clear_Rev_LocVend_IntStore_GRMultDelivNotecontTradGood
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 24th Oct
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02GR04_002_Clear_Rev_LocVend_IntStore_GRMultDelivNotecontTradGood"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR03_004_Loc_Vend_DSD_GR_Deliv_Note_w_Trading_Goods.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode MR11SHOW----------------------------

''Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Acct maint\. document","KBKP-BELNR","",DT_MR11SHOW_100_ACCT_MAINT_DOCUMENT,False)   
Call SetTextbox("Fiscal year","KBKP-GJAHR","",DT_MR11SHOW_100_FISCAL_YEAR,False)  

'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()

'Click on Reverse Account Maintenance Document Button
Call ClickButton("Reverse Account Maintenance Document   \(Shift\+F5\)",False)
Wait(5)


'Enter Posting Date
Call SetTextbox("Posting Date","MR11_HEAD-BUDAT","",ConvertDate(DT_MR11SHOW_200_POSTING_DATE_OCC2),True) 

'Click on Reverse Button
Call ClickButton("Reverse   \(Shift\+F1\)",True)
Wait(5)
'Get The Document No
Call GetGridContent("",0,"Message Text",1,"Application Area","M8","DT_MR11SHOW_500_GETCELLVALUE_OF_GRIDCELL_0_CMF_TEXT_OUTPUT")


''----------------------Tcode WE09----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MR11SHOW_500_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Wait(2)

'Fill The Details and Execute
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","","/ZMDIV_MR11",True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SetTextbox("for Value \.\.\.","VALUE1_1","",DT_MR11SHOW_1000_FOR_VALUE_,False)
Wait(2)
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",False) 

Call VerifyStatusBarMessageType("S")

Call VerifyStatusBar("IDocs were found")

'Navigate to Idoc Tree and Verify the details
Call GetLabelContentByRefLabel("IDoc number", 0, -48, "DT_IDOC_NUMBER", False)
Call ClickLabel(DT_IDOC_NUMBER,0,False)

Call ClickLinkGuiTree("IDoc "& DT_IDOC_NUMBER&";Data records;ZZMDIV_E1RBKP","ZZMDIV_E1RBKP","",False)
Call TakeScreenShot()
Call ClickLinkGuiTree("IDoc "& DT_IDOC_NUMBER&";Data records;ZZMDIV_E1RSEG","ZZMDIV_E1RSEG","",False)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()
''
'''------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
