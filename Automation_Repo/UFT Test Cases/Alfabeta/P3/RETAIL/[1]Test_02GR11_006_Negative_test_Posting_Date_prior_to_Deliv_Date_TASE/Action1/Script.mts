

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02GR11_006_Negative_test_Posting_Date_prior_to_Deliv_Date
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

gstrTestCaseName = "Test_02GR11_006_Negative_test_Posting_Date_prior_to_Deliv_Date"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'''--------------------------------MIGO-----------------------------
'''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULT_TVBWART_OCC2,False)

Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_MIGO_2000_GODYNPROPO_NUMBER,False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_MIGO_0110_DELIVERY_NOTE,False)
Call SetTextbox("Posting Date","GOHEAD-BUDAT","",ConvertDate(DT_MIGO_0110_POSTING_DATE_OCC4),False)
Call ClickButtonIfExist("MIGO_OK_GO",False)
'
Call SetTableData("SAPLMIGOTV_GOITEM","OK",1,"","",DT_MIGO_0304_ITEM_OK_OCC3,False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,DT_MIGO_0304_ITEM_OK_OCC5,False)
Call ClickButton("Check Entries   \(F7\)",False)
Date1=ConvertDoubledigit(CSTR(Day(Date))) +"."+ ConvertDoubledigit(Cstr(Month(Date))) + "."+ ConvertDoubledigit(Cstr(Year(Date)))
Date2=ConvertDoubledigit(CSTR(Day(DT_MIGO_0110_POSTING_DATE_OCC4))) +"."+ ConvertDoubledigit(Cstr(Month(DT_MIGO_0110_POSTING_DATE_OCC4))) + "."
Call VerifyifGuiLabelExists("Document date " & Date1 &" cannot be later than the Posting date "& Date2)
Call ClickButton("Continue   \(Enter\)",False)

Call LogOff()
Call FinalStatus ()
