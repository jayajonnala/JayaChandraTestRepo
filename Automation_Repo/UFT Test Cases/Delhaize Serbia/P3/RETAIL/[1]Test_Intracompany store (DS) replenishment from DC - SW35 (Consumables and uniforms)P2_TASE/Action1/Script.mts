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
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Intracompany store (DS) replenishment from DC - SW35 (Consumables and uniforms)P2
'.................Author : TCS        :Bitan
'................ Creation Date    : 7th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Intracompany store (DS) replenishment from DC - SW35 (Consumables and uniforms)P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DS\RETAIL\DT_Intracompany store (DS) replenishment from DC - SW35 (Consumables and uniforms)p2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'----------------------------------------ME23N------------------------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()  
Call TakeScreenShot()

Call ClickButtonIfExist("Other Purchase Order   \(Shift\+F5\)",False)
wait(2)
Call SetTextboxPopupIfExist("MEPO_SELECT-EBELN","Pur\. Order",DT_ME23N_0003_PUR_ORDER)
Call PressEnter()
Wait(2)
Call TakeScreenShot()

''Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
'''Call SelectTab("HEADER_DETAIL","Org. Data",False)
'''If (VerifyTextBoxEnabled("Purch\. Org\.","MEPO1222-EKORG","",False)=False) Then
'''Call ClickButtonIfExist("Display/Change \(F7\)",False)
'''End If

Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)
Wait(1)
Call TakeScreenShot()

Call GetGridContent("",0,"Article Document",1,"Short Text","Lfs","DT_PO_DELIVERY")
Call GetGridContent("",0,"Posting Date",1,"Short Text","Lfs","DT_ME23N_1105_CHECK_TEXT_OF_DOC_DATE_OUTPUT")

''Get the Article No and Document Posting Date
'Call GetGridContent("",0,"Article Document",1,"Local currency","RSD","DT_ARTICLE_DOCUMENT_NUM_OUTPUT")
'Call GetGridContent("",0,"Posting Date",1,"Local currency","RSD","DT_POSTING_DATE_OUTPUT")

Call ClickButtonIfExist("Messages   \(Shift\+F9\)",False)
wait(2)
Call TakeScreenShot()

Call GetTableCellData("SAPDV70ATC_NAST3","Output Type","1","Description","PO for Collective PO","DT_OUTPUT_TYPE_0",False)
Call GetTableCellData("SAPDV70ATC_NAST3","Output Type","2","Description","AB Mail to Stores","DT_OUTPUT_TYPE_1",False)

Call LogOff()
Call FinalStatus ()

