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



'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_AA074- Maintain Asset Masterdata- Unflag shut down indic for LVA prev month_P2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 31th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_AA074- Maintain Asset Masterdata- Unflag shut down indic for LVA prev month_P2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_AA074- Maintain Asset Masterdata- Unflag shut down indic for LVA_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=3
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations


Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''Increment the parameter / reload
'Call WriteRunTimeDataToExcelGlobalSheet ("DT_DESC_INC",(Cint(DT_DESC_INC)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode AS01----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Time-dependent",False)
'Capture the screenshot
Call TakeScreenShot()
Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_COST_CENTER,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Allocations",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Origin",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Vendor","ANLA-LIFNR","",DT_AS01_1181_VENDOR,False)
Call PressEnter() 
Call SetTextbox("WBS element","ANLA-POSNR","",DT_AS01_1182_WBS_ELEMENT,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
wait(1)
Call TakeScreenShot()
'Validate If asset is created
Call GetStatusBar("item1","DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR)
'
''
'''----------------------Tcode F-90----------------------------
'
'Enter the Tcode
Call SetTcode(DT_AS01_0105_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_AS01_0105_OKCD)
'Capture the screenshot
Call TakeScreenShot()

''Call SetTextbox("Document Date","BKPF-BLDAT","",Replace(DT_AS01_0100_DOCUMENT_DATE,"/","."),False)
Call SetTextbox("Document Date","BKPF-BLDAT","",DT_AS01_0100_DOCUMENT_DATE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_AS01_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_AS01_0100_COMPANY_CODE,False)
''Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace(DT_AS01_0100_POSTING_DATE,"/","."),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",DT_AS01_0100_POSTING_DATE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_AS01_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_AS01_0100_REFERENCE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS01_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0100_ACCOUNT,False)

'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0302_AMOUNT,False)
Call SetTextbox("Tax Amount","BSEG-WMWST","",DT_AS01_0302_TAX_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_AS01_0302_TAX_CODE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS01_0302_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0302_ACCOUNT,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_AS01_0302_TTYPE,False)

'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0305_AMOUNT,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Wait(2)
Call VerifyTextBoxContent("C","RF05A-AZSAL","",DT_AS01_0700_CHECK_TEXT_OF_C,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call PressEnter()  
Call TakeScreenShot()

'Validate If asset is created
Call GetStatusBar("item1","DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_AS01_0100_CHECK_TEXT_OF_STATUSBAR)

'
'''----------------------Tcode FB03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_AS01_0107_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_AS01_0107_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_AS01_0100_DOCUMENT_NUMBER,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
Call ClickButtonIfExist("Choose   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

'validate grid components
call VerifyGridCellContent("",1,"Account","",DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",2,"Account","",DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
call VerifyGridCellContent("",3,"Account","",DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

'
'''----------------------Tcode AS02----------------------------
'
'Enter the Tcode
Call SetTcode(DT_AS01_0750_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_AS01_0750_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS01_0100_ASSET,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS01_0100_SUBNUMBER,False)
Call PressEnter()    
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Time-dependent",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("More Intervals",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Insert Row",False)
Call SetTextbox("From-date of new interval","ANLZ-ADATU","",DT_AS01_3010_FROMDATE_OF_NEW_INTERVAL,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Yes   \(Enter\)",True)

'Capture the screenshot
Call TakeScreenShot()
Call SetTableDataNoRef("SAPLAISTTIME","Shutdown",1,DT_AS01_3000_TABLECELL_SHUTDOWN_0,False)

Call ClickButton("Save   \(Ctrl\+S\)",False)
'Validate If asset is created
VerifyStatusBar(DT_AS01_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)


Call ClickButton("Asset values   \(Ctrl\+F1\)",False) 
Call ClickButtonIfExist("Continue   \(Enter\)",True) 
'Capture the screenshot
Call TakeScreenShot()

'select node
Call SelectNodeGuiTree("0","Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")

Call SelectTab("IDC_TABSTRIP",DT_AS01_0100_POSTED_VALUES,False)
'Capture the screenshot
Call TakeScreenShot()
'get grid contents
Call GetGridContentByTitle("Depreciation posted/planned","","Depreciation period",7,"DT_AS01_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_PERAF_OUTPUT")
Call GetGridContentByTitle("Depreciation posted/planned","","Ordinary dep. TBP",7,"DT_AS01_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_NAFAZ_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Verify
Call VerifyGridCellContent("Depreciation posted/planned",7,"Depreciation period","",DT_AS01_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_PERAF)
Call VerifyGridCellContent("Depreciation posted/planned",7,"Ordinary dep. TBP","",DT_AS01_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_NAFAZ)

Call SelectTab("IDC_TABSTRIP",DT_AS01_0100_PARAMETERS,False)
'Capture the screenshot
Call TakeScreenShot()
'Verify
Call VerifyTextBoxContent("Ord\.dep\.start date","AW01_DEP_PAR-AFABG","",Replace(DT_AS01_0304_CHECK_TEXT_OF_ORDDEPSTART_DATE,"/","."),False)
Call VerifyTextBoxContent("/","AW01_DEP_PAR-NDPER","",DT_AS01_0304_CHECK_TEXT_OF_AW01_DEP_PARNDPER,False)
'Capture the screenshot
Call TakeScreenShot()


'select node
Call SelectNodeGuiTree("0","Depreciation Areas;D1 Non-Leading (Local)")

Call SelectTab("IDC_TABSTRIP",DT_AS01_0100_POSTED_VALUES_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
'get grid contents
Call GetGridContentByTitle("Depreciation posted/planned","","Depreciation period",7,"DT_AS01_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_PERAF_OCC1_OUTPUT")
Call GetGridContentByTitle("Depreciation posted/planned","","Ordinary dep. TBP",7,"DT_AS01_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_NAFAZ_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Verify
Call VerifyGridCellContent("Depreciation posted/planned",7,"Depreciation period","",DT_AS01_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_PERAF_OCC1)
Call VerifyGridCellContent("Depreciation posted/planned",7,"Ordinary dep. TBP","",DT_AS01_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_NAFAZ_OCC1)

Call SelectTab("IDC_TABSTRIP",DT_AS01_0100_PARAMETERS_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
'Verify
Call VerifyTextBoxContent("Ord\.dep\.start date","AW01_DEP_PAR-AFABG","",Replace(DT_AS01_0304_CHECK_TEXT_OF_ORDDEPSTART_DATE_OCC1,"/","."),False)
Call VerifyTextBoxContent("Dep\. Key","AW01_DEP_PAR-AFASL","",DT_DEPKEY,False)
Call VerifyTextBoxContent("/","AW01_DEP_PAR-NDPER","",DT_AS01_0304_CHECK_TEXT_OF_AW01_DEP_PARNDPER_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

