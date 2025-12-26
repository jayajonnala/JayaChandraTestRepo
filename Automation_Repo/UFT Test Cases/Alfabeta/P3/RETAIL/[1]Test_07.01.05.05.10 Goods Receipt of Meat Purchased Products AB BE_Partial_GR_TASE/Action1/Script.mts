

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_07.01.05.05.10 Goods Receipt of Meat Purchased Products AB BE_Partial_GR
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

gstrTestCaseName = "Test_07.01.05.05.10 Goods Receipt of Meat Purchased Products AB BE_Partial_GR"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
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
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''--------------------------------MIGO-----------------------------

call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC_OCC1)
Call TakeScreenshot()
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MOVEMENT_TYPE,False)
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_MIGO_2000_GODYNPROPO_NUMBER,False)
Call TakeScreenshot()
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_MIGO_0110_DELIVERY_NOTE,False)
Call TakeScreenshot()
Call PressEnter()    ' 
Call ClickButtonIfExist("Open detail data",False)
Call SelectTab("TS_GOITEM","Quantity",False)
Call TakeScreenshot()
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,DT_MIGO_0304_ITEM_OK,False)
Call SetTextbox("Qty in Unit of Entry","GOITEM-ERFME","",DT_MIGO_0315_QTY_IN_UNIT_OF_ENTRY,False)
Call TakeScreenshot()
Call SetTextbox("Qty in Unit of Entry","GOITEM-ERFMG","",DT_QTY_TO_GR,False)
Call PressEnter()     ' 

Call TakeScreenShot()
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_MIGO_DOC_OUTPUT")
Call VerifyStatusBar("Article document "& DT_MIGO_DOC_OUTPUT &" posted")


Call LogOff()
Call FinalStatus ()


