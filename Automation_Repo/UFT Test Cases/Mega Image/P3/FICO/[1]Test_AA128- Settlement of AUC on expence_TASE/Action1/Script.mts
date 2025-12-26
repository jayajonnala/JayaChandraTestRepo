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
'.................Test Script Name : Test_AA128- Settlement of AUC on expence_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 31th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_AA128- Settlement of AUC on expence_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_AA128- Settlement of AUC on expence_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'Increment the parameter / reload
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''----------------------Tcode AS01----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
Call SetTextbox("Number of similar assets","RA02S-NASSETS","",DT_AS01_0105_NUMBER_OF_SIMILAR_ASSETS,False)
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
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

''----------------------Tcode F-90----------------------------

'Enter the Tcode
Call SetTcode(DT_AS01_0105_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_AS01_0105_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",Replace(DT_AS01_0100_DOCUMENT_DATE,"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_AS01_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_AS01_0100_COMPANY_CODE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_AS01_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_AS01_0100_REFERENCE,False)
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

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call PressEnter()  
Call TakeScreenShot()
'Validate If asset is created
Call GetStatusBar("item1","DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_AS01_0100_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
'
'''----------------------Tcode FB03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_AS01_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_AS01_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR,False)
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
'''----------------------Tcode AIAB----------------------------
'
'Enter the Tcode
Call SetTcode(DT_AS01_0750_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_AS01_0750_OKCD)
'Capture the screenshot
Call TakeScreenShot()
''''
'''''Call SetTextbox("Company Code","AICOM-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
'''''Call SetTextbox("Asset","AICOM-ANLN1","",DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR,False)
'''''Call SetTextbox("Sub-number","AICOM-ANLN2","",0,False)
''''''Capture the screenshot
'''''Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRowGuiGridbyRowNo("","",DT_AS01_0500_GRIDCELL_0__OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Enter distribution rules   \(Shift\+F6\)",False)
'Capture the screenshot
Call TakeScreenShot()

'enter table data
Call SetTableDataNoRef("SAPLKOBSTC_RULES","Cat",1,DT_AS01_0130_TABLECELL_CAT_0,FALSE)
Call SetTableDataNoRef("SAPLKOBSTC_RULES","Settlement Receiver",1,DT_AS01_0130_TABLECELL_SETTLEMENT_RECEIVER_0,FALSE)
'Call SetTableDataNoRef("SAPLKOBSTC_RULES","%",1,DT_AIAB_0130_TABLECELL__0,FALSE)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()
'check if data saved successfully
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)
VerifyStatusBar(DT_AS01_0500_CHECK_TEXT_OF_STATUSBAR)
'
'''----------------------Tcode AIBU----------------------------
'
'Enter the Tcode
Call SetTcode(DT_AS01_0500_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_AS01_0500_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset Val\. Date","ANEP-BZDAT","",Replace(DT_AS01_0100_ASSET_VAL_DATE,"/","."),False)
Call SetTextbox("Document type","\*KOMK1-BLART","",DT_AS01_0100_DOCUMENT_TYPE,False)
'Call SetTextbox("Company code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
'Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR,False)
'Call SetTextbox("Sub-number","ANLA-ANLN2","",0,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
Call ClickButton("Back   \(F3\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckbox("LKO74-TESTLAUF",1,DT_AS01_0100_TEST_RUN,False)
Call ClickButton("Execute   \(F8\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

'Validate If doc is created
Call GetStatusBar("item1","DT_AS01_0500_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_AS01_0500_CHECK_TEXT_OF_STATUSBAR_OCC1)
Call VerifyStatusBarMessageType("S")

'
'''----------------------Tcode FB03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_AS01_0500_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_AS01_0500_OKCD_OCC1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_AS01_0500_CHECK_MESSAGEPARAMETER_OF_STATUSBAR,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
Call ClickButtonIfExist("Choose   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

'validate grid components
call VerifyGridCellContent("",1,"Posting Key","",DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",2,"Posting Key","",DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

call VerifyGridCellContent("",1,"Account","",DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
call VerifyGridCellContent("",2,"Account","",DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)

call VerifyGridCellContent("",1,"Amount","",DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
call VerifyGridCellContent("",2,"Amount","",DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)

'following tow is disabled as these fields are not present in default layout
'call VerifyGridCellContent("",1,"Cost Center","","")
'call VerifyGridCellContent("",2,"Cost Center","",DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOSTL)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


