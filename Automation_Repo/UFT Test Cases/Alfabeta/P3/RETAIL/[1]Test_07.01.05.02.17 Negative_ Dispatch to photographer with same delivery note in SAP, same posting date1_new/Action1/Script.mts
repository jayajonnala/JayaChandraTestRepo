	

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_POST_DeleteVAT_from_Customer
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'gstrTestCaseName = "Test_07.01.05.02.17 Negative_ Dispatch to photographer with same delivery note in SAP, same posting date"
gstrTestCaseName = "Test_07.01.05.02.17 Negative_ Dispatch"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_07.01.05.02.17 Negative_ Dispatch to photographer with same delivery note in SAP, same posting date_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------MIGO -----------------------------------
' SelectRadioButton(radiobuttonName, radiobuttonAttachedtext, blnIsItPopup)

Call SetCombo("GODYNPRO-ACTION","Transfer Posting")
Call SetCombo("GODYNPRO-REFDOC","Other")
Call PressEnter()

Call SetTextbox("Article Slip","GOHEAD-MTSNR","",DT_MIGO_0112_ARTICLE_SLIP,False) 

Call ClickButtonIfExist("Open detail data",False)
Call PressEnter()

Call SetTextbox("Stor\. Loc\.","GODYNPRO-LGOBE","",DT_MIGO_0390_STOR_LOC, False)
Call SetTextbox("Site","GODYNPRO-NAME1","",DT_MIGO_0390_SITE,False)
Call SetTextbox("Article","GODYNPRO-MAKTX","",DT_MIGO_0390_ARTICLE,False)
Call SetTextbox("Vendor Transfer Pstg","GOITEM-UMMAT_VENDORNAME","",DT_MIGO_0396_GOITEMUMMAT_VENDORNAME, False)
Call SetTextbox("Qty in UnE","GODYNPRO-ERFMG","",DT_MIGO_0390_QTY_IN_UNE,False)
Call PressEnter()

Call GetTextboxValue("GOHEAD-BLDAT","","DT_MIGO_0001_GET_DATE",False)

Call WriteRunTimeDataToExcelGlobalSheet("DT_MIGO_0001_GET_DATE",DT_MIGO_0001_GET_DATE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call VerifyifGuiLabelExists(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus ()


