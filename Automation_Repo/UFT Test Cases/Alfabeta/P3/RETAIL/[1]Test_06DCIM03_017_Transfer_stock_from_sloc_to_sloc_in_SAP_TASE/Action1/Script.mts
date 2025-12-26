

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06DCIM03_017_Transfer_stock_from_sloc_to_sloc_in_SAP
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

gstrTestCaseName = "Test_06DCIM03_017_Transfersloctosloc_SAP"
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
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''--------------------------------MIGO-----------------------------


'
call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
Call TakeScreenShot()

Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call TakeScreenShot()


Call SetTextbox("Document Date","GOHEAD-BLDAT","",ConvertDate(DT_MIGO_0112_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","GOHEAD-BUDAT","",ConvertDate(DT_MIGO_0112_POSTING_DATE),False)
Call TakeScreenShot()
Call ClickButtonIfExist("Open detail data",False)

 Call SetTextbox("Qty in UnE","GODYNPRO-ERFMG","",DT_MIGO_0390_QTY_IN_UNE,False)
Call SetTextboxNoLabel("GODYNPRO-ERFME",0,DT_MIGO_0390_QTY_IN_UNE_OCC1,False)
 Call SetTextbox("Article","GODYNPRO-MAKTX","",DT_MIGO_0390_ARTICLE,False)
 Call SetTextbox("Site","GODYNPRO-NAME1","",DT_MIGO_0390_SITE,False)
 Call SetTextbox("Stor\. Loc\.","GODYNPRO-LGOBE","",DT_MIGO_0390_STOR_LOC,False)
Call SetTextboxNoLabel("GOITEM-UMLGOBE",0,DT_MIGO_0390_GOITEMUMLGOBE,False)

Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

Call ClickButtonIfExist("Check Entries   \(F7\)",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)  
Call PressEnter()     ' 
'Call VerifyStatusBarMessageType("S")
Call TakeScreenShot()
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  


Call GetStatusBar("item1","DT_ARTICLE_DOC_OUTPUT")
Call VerifyStatusBar("Article document "& DT_ARTICLE_DOC_OUTPUT &" posted")

Call LogOff()
Call FinalStatus ()


