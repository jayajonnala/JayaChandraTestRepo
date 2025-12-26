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

 '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Customer open items clearing FIFO method_p4_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 3rd May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Customer open items clearing FIFO method_p4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Customer open items clearing  FIFO method_p4_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode FBL5N----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("X_CLSEL",DT_FBL5N_1000_CLEARED_ITEMS,False)
Call SetTextbox("Clearing date","SO_AUGDT-LOW","",ConvertDateFormat(DT_FBL5N_1000_CLEARING_DATE),False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FBL5N_1000_COMPANY_CODE,False)
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FBL5N_1000_CUSTOMER_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()


Call ClickButton("Execute   \(F8\)",False)
Wait(5)
'
'''''''''''''Additional code to filter out the correct details as layout got changed to "grid type fro label format'''''''''''
'''''''Call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",False)
'''''''Call SelectRowGuiGridbyRowNo("Column Set","",5,True)
'''''''
'''''''Call ClickButton("Show Selected Fields \(F7\)",True)
'''''''Call ClickButton("Transfer   \(Enter\)",True)
''''''
'''''''Call SelectColumnGuiGrid("","","Posting Key",True)
'''''''Call SelectRowGuiTable("SAPLSKBHTC_FIELD_LIST","Content","Posting Key",True)
'''''''Call ClickButton("Sort in descending order   \(Ctrl\+Shift\+F4\)",False)
'''''''Call GetGridContentByTitle("","","Document Number",1,"DOC_NO_OUTPUT")
'''''''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''''''Call SelectColumnGuiGrid("","","Document Number",False)
'''''''Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)
'''''''Call SetTextbox("Document Number","%%DYN001-LOW","",DOC_NO,True)
'''''''Call ClickButton("Execute   \(Enter\)",True)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''Wait(2)
'''''''Capture the screenshot
'''''''Call TakeScreenShot()
Call ClickButton("Select layout\.\.\.   \(Ctrl\+F9\)",False)

Call ClickLabel("AUTO",0,True)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call SelectRowGuiTable("SAPLSKBHTC_FIELD_LIST_820","Content","Document type",True)
Call ClickButton("Show sel\. fields \(CTRL\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)

Call SetTextbox("Document type","%%DYN001-LOW","","CL",True)
Call ClickButton("Execute   \(Enter\)",True)

'verify the details

'''''Call VerifyGridCellContent("",1,"Document type","",DT_Doc_Type)
'''''Call VerifyGridCellContent("",2,"Document type","",DT_Doc_Type)
'''''
'''''
'''''Call VerifyGridCellContent("",1,"Document Date","",Replace((DT_Check_Doc_Date),"/","."))
'''''Call VerifyGridCellContent("",2,"Document Date","",Replace((DT_Check_Doc_Date),"/","."))
'''''
'''''Call VerifyGridCellContent("",1,"Posting Key","",DT_FBL5N_0120_CHECK_TEXT_OF_NO_NAME_OCC4)
'''''Call VerifyGridCellContent("",2,"Posting Key","",DT_FBL5N_0120_CHECK_TEXT_OF_NO_NAME_OCC6)
'''''
'''''Call VerifyGridCellContent("",3,"Amount in local currency","",DT_Balance)


Call VerifyifGuiLabelExists_ByIndex(DT_Doc_Type,0)
Call VerifyifGuiLabelExists_ByIndex(DT_Doc_Type,1)
Call VerifyifGuiLabelExists_ByIndex(Replace((DT_Check_Doc_Date),"/","."),0)
Call VerifyifGuiLabelExists_ByIndex(Replace((DT_Check_Doc_Date),"/","."),1)
Call VerifyifGuiLabelExists_ByIndex(DT_FBL5N_0120_CHECK_TEXT_OF_NO_NAME_OCC4,0)
Call VerifyifGuiLabelExists_ByIndex(DT_FBL5N_0120_CHECK_TEXT_OF_NO_NAME_OCC4,1)
Call VerifyifGuiLabelExists_ByIndex(DT_Balance,1)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

