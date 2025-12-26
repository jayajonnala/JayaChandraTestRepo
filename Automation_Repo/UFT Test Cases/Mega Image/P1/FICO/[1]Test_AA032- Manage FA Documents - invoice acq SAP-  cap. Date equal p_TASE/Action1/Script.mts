
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA032- Manage FA Documents - invoice acq SAP-  cap. Date equal p
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


gstrTestCaseName = "Test_AA032p"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA032- Manage FA Documents - invoice acq SAP-  cap. Date equal p.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

Call StartDateof445PeriodByDate(DT_TODAY,"DT_STARTING_DATE_PERIOD")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
''
'''''''--------------login----------------'''''
'' 
'
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ASSET_SUBNUMBER",Cint(DT_ASSET_SUBNUMBER)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''''--------TransactionCode-ZMDPU_INFOREC_COPY----------''''

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS11_0110_ASSET,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS11_0110_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Description","ANLA-TXT50","",DT_AS11_1140_DESCRIPTION,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Cost Center","ANLZ-KOSTL","",DT_AS11_1145_CHECK_TEXT_OF_COST_CENTER,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Allocations", False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Origin", False)
Call SetTextbox("Vendor","ANLA-LIFNR","",DT_AS11_1181_VENDOR,False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("WBS element","ANLA-POSNR","",DT_AS11_1182_WBS_ELEMENT,False)
Call TakeScreenShot
Call PressEnter()
Call SelectTab("TABSTRIP100", "Leasing", False)
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Deprec. Areas", False)
Call VerifyTableCellContent(1, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_0)
Call VerifyTableCellContent(2, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_1)
Call VerifyTableCellContent(3, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_2)
Call VerifyTableCellContent(4, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_3)
Call VerifyTableCellContent(5, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_4)
Call TakeScreenShot
Call PressEnter()
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Create",True)
''Call GetStatusBar("item1", "DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item1", "DT_AS11_0110_OUTPUT")

''''''''--------TransactionCode-/NF-90----------''''
'
Call SetTcode(DT_AS11_0110_OKCD) 
Call TakeScreenShot
Call PressEnter()  

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call SetTextbox("Period","BKPF-MONAT","",DT_AS11_0100_PERIOD,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_AS11_0100_DOCUMENT_DATE),False)
Call SetTextbox("Type","BKPF-BLART","",DT_AS11_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_AS11_0100_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_AS11_0100_POSTING_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_AS11_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_AS11_0100_REFERENCE,False)
Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_AS11_0100_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS11_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS11_0302_AMOUNT,False)
Call SetTextbox("Tax Amount","BSEG-WMWST","",DT_AS11_0302_TAX_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_AS11_0302_TAX_CODE,False)
Call SetTextbox("Bus. Area","BSEG-GSBER","",DT_AS11_0302_BUS_AREA,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_AS11_0302_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS11_0302_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS11_0302_ACCOUNT,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_AS11_0302_TTYPE,False)
Call PressEnter()
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS11_0305_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_AS11_0305_TEXT,False)
Call ClickButton("Display Document Overview   \(Shift\+F2\)", False)
Call PressEnter()
Call VerifyTextBoxContent("C", "RF05A-AZSAL", "", DT_AS11_0700_CHECK_TEXT_OF_C, False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)", False)
Call TakeScreenSHot()
Call PressEnter()
Call TakeScreenShot
''Call GetStatusBar("item1", "DT_OUTPUT")
''Call GetStatusBar("item1", "DT_DOC_NUM_OUTPUT")
Call GetStatusBar("item1","DT_DOC_NUM_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NUM_OUTPUT&" was posted in company code RO02")
Call TakeScreenShot

Call ClickButton("Exit   \(Shift\+F3\)", False)
Call ClickButtonIfExist("Yes", True)


Call LogOff()

Call FinalStatus ()






'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [8,3437477]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//




