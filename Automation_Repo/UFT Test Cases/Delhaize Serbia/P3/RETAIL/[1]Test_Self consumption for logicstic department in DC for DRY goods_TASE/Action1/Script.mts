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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Self consumption for logicstic department in DC for DRY goods_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 14th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Self consumption for logicstic department in DC for DRY goods_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Self consumption for logicstic department in DC for DRY goods_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
''
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  


'''----------------------Tcode MIGO----------------------------


Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call WriteRunTimeDataToExcelGlobalSheet ("DT_PO_INC",(Cint(DT_PO_INC)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
Call TakeScreenShot()

'Call SetTextbox("TF trfr plnt to plnt","GODEFAULT_TV-BWART","",DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call SetTextboxNoLabel("GODEFAULT_TV-BWART","",DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call TakeScreenShot()
'Call PressEnter()  

Call SetTextbox("Article","GODYNPRO-MAKTX","",DT_MIGO_0390_ARTICLE,False)
Call SetTextbox("Qty in UnE","GODYNPRO-ERFMG","",DT_MIGO_0390_QTY_IN_UNE,False)
Call SetTextbox("Site","GODYNPRO-NAME1","",DT_MIGO_0390_SITE,False)
Call SetTextbox("Stor\. Loc\.","GODYNPRO-LGOBE","",DT_MIGO_0390_STOR_LOC,False)
Call SetTextbox("SLoc Transfer Pstg","GOITEM-UMLGOBE","",DT_MIGO_0390_GOITEMUMLGOBE,False)

Call ClickButtonIfExist("Cancel",True) 
'Call SetTextbox("Site Trfr Pstg","GOITEM-UMNAME1","",DT_MIGO_0390_SITE,False)
'Call SetTextbox("Plant Trfr Pstg","GOITEM-UMNAME1","",DT_MIGO_0390_SITE,False)''not in log but it mandates now (only for movement type 301)

Call PressEnter()
Call TakeScreenShot()

Call SetTextbox("Document Date","GOHEAD-BLDAT","",Replace(DT_MIGO_0112_DOCUMENT_DATE,"/","."),False)
Call SetTextbox("Posting Date","GOHEAD-BUDAT","",Replace(DT_MIGO_0112_POSTING_DATE,"/","."),False)
Call PressEnter() 
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Wait(2)
Call ClickButtonIfExist("Save",True) 
Wait(2)
Call TakeScreenShot()

Call GetStatusBar("item1","DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot()
'
'
'''----------------------Tcode VA01----------------------------

Call SetTcode(DT_MIGO_0001_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MIGO_0001_OKCD)
Call TakeScreenShot()

Call SetTextbox("Sales Organization","VBAK-VKORG","",DT_MIGO_0101_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","VBAK-VTWEG","",DT_MIGO_0101_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Division","VBAK-SPART","",DT_MIGO_0101_DIVISION,False)
Call SetTextbox("Order Type","VBAK-AUART","",DT_MIGO_0101_ORDER_TYPE,False)
Call TakeScreenShot()

Call PressEnter() 
Call TakeScreenShot()

Call SetTextbox("Purch\. Order No\.","VBKD-BSTKD","",DT_MIGO_4021_PO_NUMBER,False)
Call SetTextbox("Sold-To Party","KUAGV-KUNNR","",DT_MIGO_4701_SOLDTO_PARTY,False)
Call SetTextbox("Ship-To Party","KUWEV-KUNNR","",DT_MIGO_4701_SHIPTO_PARTY,False)
Call SetTextbox("Payment terms","VBKD-ZTERM","",DT_MIGO_4440_PAYMENT_TERMS,False)
Call SetTextbox("Incoterms","VBKD-INCO1","",DT_MIGO_4440_INCOTERMS,False)
Call SetTextbox("Incoterms","VBKD-INCO2","",DT_MIGO_4440_INCOTERMS_OCC1,False)
Call SetComboByKey("VBAK-AUGRU",DT_MIGO_4440_ORDER_REASON)
Call TakeScreenShot()

Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG","Article",1,DT_MIGO_4900_TABLECELL_ARTICLE_0,False)
Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG","Order Quantity",1,DT_MIGO_4900_TABLECELL_ORDER_QUANTITY_0,False)
Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG","Site",1,DT_MIGO_4900_TABLECELL_SITE_0,False)

Call PressEnter() 
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(2)
Call TakeScreenShot()

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Wait(2)
Call TakeScreenShot()

Call ClickButtonIfExist("Edit",True)
Wait(2)
Call TakeScreenShot()

Call DoubleClick()
Wait(2)
Call SetTextbox("Doc\. Currency","VBAK-WAERK","",DT_CURRENCY,False)
Call PressEnter() 
Call TakeScreenShot()

Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)
Call TakeScreenShot()

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Wait(2)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_MIGO_4001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item2","DT_MIGO_4001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_MIGO_4001_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType("S")

Call LogOff()
Call FinalStatus ()


