'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Distribution of consumables_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 26th Feb
'.................Modified By :
'.................Modified Date/Details :
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

gstrTestCaseName = "Test_Distribution of consumables_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Distribution of consumables_p2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'''----------------------Tcode ME23N----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()
'Enter the Details
Call ClickButtonIfExist("Other Purchase Order   \(Shift\+F5\)",False)
'Call SetTextboxPopupIfExist("MEPO_SELECT-EBELN","Pur\. Order",DT_ME23N_0003_PUR_ORDER)
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PUR_ORDER,True)
Call TakeScreenShot()
Call ClickButton("Other Document   \(Enter\)",True)

'close document overview if exist
Call ClickButtonIfExist("Switch Off Document Overview   \(F9\)",False)

'Close help page if exist
If SAPGuiSession("transaction:="&DT_SAPTRANSACTIONCODE).SAPGuiWindow("transaction:="&DT_SAPTRANSACTIONCODE).Page("title:=").Exist = True Then	
	Call ClickButton("Help   \(Shift\+F1\)",False)
	Wait(1)	
End If

'navigate to item details section
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
'Texts tab
Call SelectTab("ITEM_DETAIL","Texts",False)
Call TakeScreenShot()
'Purchase order history tab
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)
Call TakeScreenShot()
'Call ClickButtonToolBar("&FIND","")
Call ClickButtonToolBar("&FIND",0)
'Call SetTextboxPopupIfExist("GS_SEARCH-VALUE","Search Term:",DT_ME23N_0841_SEARCH_TERM) 
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ME23N_0841_SEARCH_TERM,True)
Call ClickButtonIfExist("OK   \(Enter\)",True)
Wait(1)
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonToolBar("&DETAIL",0)
Wait(1)
Call TakeScreenShot()
'get article document nummber
Call GetCellDataGuiGridPopupByRefTwoColumns("","","Cell Content","Group Description","Article Document","Group Description","Article Document","DT_ME23N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_VALUE_OUTPUT")
Call ClickButtonIfExist("Close window   \(Enter\)",True)
'reload data sheet
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''----------------------Tcode VL03N----------------------------
'Enter the Tcode
Call SetTcode(DT_ME23N_0014_OKCD) 
Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()
'Enter data
Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_ME23N_4004_OUTBOUND_DELIVERY,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()  
Call SelectMenuBar("Extras;Delivery Output;Header")
Call VerifyTableCellContent(1,"Output Type","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Display <-> Change   \(Ctrl\+F1\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Post Goods Issue   \(Shift\+F8\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
Call TakeScreenShot()

Call GetStatusBar("item3","DT_ARTICLE_DOCUMENT_NUM_OUTPUT")
Wait(2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyStatusBar("Outb. del.(Affiilia) "&DT_ME23N_4004_OUTBOUND_DELIVERY&" saved, article document "&DT_VL02N_0200_GRIDCELL_0_ART_DOC_NUMBER&" created")


'Call VerifyStatusBar(DT_ME23N_4004_CHECK_TEXT_OF_STATUSBAR)
'-----------------------------------------------Document flow-----------------------------------
Call ClickButton("Document Flow   \(F7\)",False)
Wait(1)
Call ClickButtonToolBar("&FIND",0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ME23N_0840_CELL_CONTENT,True)
Call ClickButtonIfExist("OK   \(Enter\)",True)
Wait(1)
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call TakeScreenShot()
Call ClickButton("Display document   \(F8\)",False)
Wait(1)
'''Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
Call TakeScreenShot()
Call ClickButton("Accounting Documents   \(F7\)",False)
Wait(1)
Call SelectCellGuiGrid("Documents in Accounting","",DT_ME23N_0200_GRIDCELL_0_DOC_NUMBER,"Doc. Number",true)
Call SelectRowGuiGridbyRowNo("Documents in Accounting","",DT_ME23N_0200_GRIDCELL_0_DOC_NUMBER,true)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Display Document   \(F2\)",True)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()
Call GetTextboxValue("BKPF-BELNR","","DT_ME23N_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",False)
Call VerifyTextBoxContent("Reference","BKPF-XBLNR","",DT_ME23N_0750_CHECK_TEXT_OF_REFERENCE,False)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

