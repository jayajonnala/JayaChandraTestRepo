

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.12.02 Maintain Asset Value_Asset acquisition_Unique_P2
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrTestCaseName = "Test_09.11.01.12.02 Maintain Asset Value_Asset acquisition_Unique_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR10_002_LocVend_w_Subrange_DSD_GR_Deliv_Note_w_Trading_Goods.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''----------------------Login----------------------------
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  


'----------------------Tcode AS01----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))

'Enter details
Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
Call SetTextbox("Number of similar assets","RA02S-NASSETS","",DT_AS01_0105_NUMBER_OF_SIMILAR_ASSETS,False)
Call SetTextbox("Asset","RA02S-RANL1","",DT_AS01_0105_ASSET,False)
Call SetTextbox("Sub-number","RA02S-RANL2","",DT_AS01_0105_SUBNUMBER,False)
Call SetTextbox("Company code","RA02S-RBUKR","",DT_AS01_0105_COMPANY_CODE_OCC1,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",False)

Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

Call GetTextboxValue("ANLA-TXT50", "", "DT_AS01_1140_CHECK_TEXT_OF_DESCRIPTION_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_AS01_1140_CHECK_TEXT_OF_DESCRIPTION_OUTPUT",DT_AS01_1140_CHECK_TEXT_OF_DESCRIPTION)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call TakeScreenShot()

'Navigate to Time-dependent Tab
Call SelectTab("TABSTRIP100","Time-dependent",False)

Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_PROFIT_CENTER,False)
Call SetTextbox("Segment","ANLZ-SEGMENT","",DT_AS01_1145_SEGMENT,False)
Call TakeScreenShot()

'Navigate to Allocation Tab
Call SelectTab("TABSTRIP100","Allocations",False)
Call TakeScreenShot()

'Navigate to Origin Tab
Call SelectTab("TABSTRIP100","Origin",False)
Call TakeScreenShot()

'Navigate to Leasing Tab
Call SelectTab("TABSTRIP100","Leasing",False)
Call TakeScreenShot()

'Navigate to Deprec. Areas Tab
Call SelectTab("TABSTRIP100","Deprec. Areas",False)
Call TakeScreenShot()

'Post the Document
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True) 
Wait(2)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call GetStatusBar("item1","DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("The asset "&DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" 0 is created" )
Call WriteRunTimeDataToExcelGlobalSheet("DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''''----------------------Tcode ABSO----------------------------
Call SetTcode(DT_AS01_0105_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Company Code","ANBZ-BUKRS","",DT_AS01_0100_COMPANY_CODE,False)
Call SetTextbox("Asset","ANBZ-ANLN1","",DT_AS01_0100_ASSET,False)
Call SetTextbox("Sub-number","ANBZ-ANLN2","",DT_AS01_0100_SUBNUMBER,False)
Call SetTextbox("Document Date","ANEK-BLDAT","",ConvertDate(DT_AS01_0100_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","ANEK-BUDAT","",ConvertDate(DT_AS01_0100_POSTING_DATE),False)
Call SetTextbox("Posting period","ANBZ-PERID","",DT_AS01_0100_POSTING_PERIOD,False)
Call SetTextbox("Transaction Type","ANBZ-BWASL","",DT_AS01_0100_TRANSACTION_TYPE,False)
Call TakeScreenShot()
Call PressEnter()  
Call TakeScreenShot()

Call SetTextbox("Amount posted","ANBZ-DMBTR","",DT_AS01_0110_AMOUNT_POSTED,False)
Call SetTextbox("Text","ANEK-SGTXT","",DT_AS01_0110_TEXT,False)
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()

Call GetTextboxValue("ANEK-SGTXT", "", "DT_AS01_0110_CHECK_TEXT_OF_TEXT_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_AS01_0110_CHECK_TEXT_OF_TEXT_OUTPUT",DT_AS01_0110_CHECK_TEXT_OF_TEXT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call TakeScreenShot()
Call ClickButton("Line Items   \(Shift\+F1\)",False)
Call TakeScreenShot()

Call VerifyTableCellContent(1, "Area", "SAPMA01BTCTRL_ANBTR", DT_AS01_0250_CHECK_TEXT_OF_TABLECELL_AREA_0)
Call VerifyTableCellContent(1, "Tra", "SAPMA01BTCTRL_ANBTR", DT_AS01_0250_CHECK_TEXT_OF_TABLECELL_TRA_0)
Call VerifyTableCellContent(1, "Amount posted", "SAPMA01BTCTRL_ANBTR", DT_AS01_0250_CHECK_TEXT_OF_TABLECELL_AMOUNT_POSTED_0)

Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Asset transaction was posted with AA document number "&DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
Call WriteRunTimeDataToExcelGlobalSheet("DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickButton("Exit   \(Shift\+F3\)",False)


''------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************




