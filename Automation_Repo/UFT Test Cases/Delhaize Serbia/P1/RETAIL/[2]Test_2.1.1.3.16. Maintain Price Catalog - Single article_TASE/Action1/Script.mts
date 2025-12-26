
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.1.1.3.16. Maintain Price Catalog - Single article
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

gstrTestCaseName = "Test_2.1.1.3.16. Maintain Price Catalog - Single article_TASE"
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
''----------------------Tcode W_SYNC----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()
Call SetTextbox("Profile","P_PROFIL","",DT_W_SYNC_0500_PROFILE,False)
Call PressEnter()  
Call TakeScreenShot()

Call SetTextbox("Supplier","PI_04-LOW", "", DT_W_SYNC_0500_VENDOR,False)
''Call SetTextboxNoLabel("Suppli","PI_03-LOW","",DT_W_SYNC_0500_VENDOR,False)
Call SetTextbox("GTIN Base UoM","PI_06-LOW","",DT_W_SYNC_0500_GTIN_BASE_UOM,False)
Call SetTextbox("Transfer Date","PI_10-LOW","","",False)
Call FocusTextBox("GTIN Base UoM","PI_06-LOW", False)
Call TakeScreenShot()

Call ClickButton("Select   \(F8\)",False)
Call TakeScreenShot()

Call SelectNodeGuiTree(0,"#1;#1;#1")
Call TakeScreenShot()

Call ClickButton("Link to Price Catalog Maintenance   \(F7\)",False)
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call SelectRowGuiGridbyRowNo("", "", 4, false)
Call TakeScreenShot()
Call ClickButtonToolBar("CHANGE_MERCHANDISE_GROUP",False)
Wait 5
Call TakeScreenShot()

Call SetTextbox("MC ref\. artl","PRICAT_K003-REF_ARTICLE",0,DT_W_SYNC_4000_MC_REF_ARTCL,True)
If DT_W_SYNC_4000_MDSE_CAT <> Empty Then

	Call SetTextbox("Mdse Cat\.","T023-MATKL",0,DT_W_SYNC_4000_MDSE_CAT,True)
Else
	Call SetTextbox("Mdse Cat\.","T023-MATKL",0,"",True)
End If


Call TakeScreenShot()
Call ClickButton("Copy Values   \(F8\)",True)
Call TakeScreenShot()
Call ClickButton("Y",True)
Call TakeScreenShot()
Call SelectRowGuiGridbyRowNo("", "", 4, false)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Wait 5
Call SelectRowGuiGridbyRowNo("", "", 4, false)
Call ClickButton("Create/Change Article - Immediately   \(Shift\+F9\)",False)
Call TakeScreenShot()
Wait 5
Call SelectRowGuiGridbyRowNo("", "", 4, false)
Call ClickButton("Create/Change Article - Immediately   \(Shift\+F9\)",False)
Call TakeScreenShot()
Wait 5
Call SelectRowGuiGridbyRowNo("", "", 4, false)

Call ClickButtonToolBar("&REFRESH",False)
Call TakeScreenShot()
Wait 5
''''Some time after clicking on refresh the article is not getting generated, So doing refresh again for two times
Call ClickButtonToolBar("&REFRESH",False)
Call TakeScreenShot()
Wait 5

Call ClickButtonToolBar("&REFRESH",False)
Call TakeScreenShot()
Wait 5

Call GetGridContentByTitle("", 0, "Article", 5, "DT_W_SYNC_0400_CHECK_GETCELLVALUE_OF_GRIDCELL_4_MATNR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_W_SYNC_0400_CHECK_GETCELLVALUE_OF_GRIDCELL_4_MATNR_OUTPUT",DT_W_SYNC_0400_CHECK_GETCELLVALUE_OF_GRIDCELL_4_MATNR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


Call ClickCellGuiGrid("", 0, "Article", 5, "", "", False)
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

