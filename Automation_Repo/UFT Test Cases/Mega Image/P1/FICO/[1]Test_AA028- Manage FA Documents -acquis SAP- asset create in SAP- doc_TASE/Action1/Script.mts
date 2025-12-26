
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA028- Manage FA Documents -acquis SAP- asset create in SAP- doc
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


gstrTestCaseName = "Test_AA028- Manage FA Documents -acquis SAP- asset create in SAP- doc"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA028- Manage FA Documents -acquis SAP- asset create in SAP- doc.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'
Call StartDateof445PeriodByDate(DT_TODAY,"DT_STARTING_DATE_PERIOD")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''''--------------login----------------'''''

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''--------TransactionCode-ZMDPU_INFOREC_COPY----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
Call SetTextbox("Number of similar assets","RA02S-NASSETS","",DT_AS01_0105_NUMBER_OF_SIMILAR_ASSETS,False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call SetTextbox("Business Area","ANLZ-GSBER","",DT_AS01_1145_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_COST_CENTER,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Allocations", False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Evaluation group 1","ANLA-ORD41","",DT_AS01_1160_EVALUATION_GROUP_1,False)
Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Origin", False)
Call SetTextbox("Vendor","ANLA-LIFNR","",DT_AS01_1181_VENDOR,False)
Call PressEnter()
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("WBS element","ANLA-POSNR","",DT_AS01_1182_WBS_ELEMENT,False)
Call TakeScreenShot
Call PressEnter()
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call SelectCheckbox("ANLZ-XSTIL", 0, DT_AS01_1145_ASSET_SHUTDOWN, False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Deprec. Areas", False)
Call SetTableData("SAPLAISTTC_ANLB", "DKey", 1, "", "", DT_AS01_DKEY, False)
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Create",True)
Call TakeScreenShot
Call GetStatusBar("MessageType", "DT_EXPECTEDVALUE_OCC1_OUTPUT")
Call GetStatusBar("item2", "DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call GetStatusBar("item3", "DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT",DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(Lcase(DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR))

''''''--------TransactionCode-/NF-90----------''''

Call SetTcode(DT_AS01_0105_OKCD) 
Call TakeScreenShot
Call PressEnter()     
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Document Date","BKPF-BLDAT","",(DT_AS01_0100_DOCUMENT_DATE),False)
Call SetTextbox("Type","BKPF-BLART","",DT_AS01_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_AS01_0100_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",(DT_AS01_0100_POSTING_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_AS01_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_AS01_0100_REFERENCE,False)
Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_AS01_0100_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0302_AMOUNT,False)
Call PressEnter()
Call PressEnter()
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call SetTextbox("Tax Amount","BSEG-WMWST","",DT_AS01_0302_TAX_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_AS01_0302_TAX_CODE,False)
Call SetTextbox("Bus. Area","BSEG-GSBER","",DT_AS01_0302_BUS_AREA,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_AS01_0302_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS01_0302_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0302_ACCOUNT,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_AS01_0302_TTYPE,False)
Call PressEnter()
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0305_AMOUNT,False)
Call PressEnter()
Call PressEnter()
Call SetTextbox("Text","BSEG-SGTXT","",DT_AS01_0305_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS01_0305_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0305_ACCOUNT,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_AS01_0305_TTYPE,False)
Call PressEnter()
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0305_AMOUNT_OCC1,False)
Call PressEnter()
Call PressEnter()
Call SetTextbox("Text","BSEG-SGTXT","",DT_AS01_0305_TEXT_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS01_0305_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0305_ACCOUNT_OCC1,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_AS01_0305_TTYPE_OCC1,False)
Call PressEnter()
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0305_AMOUNT_OCC2,False)
Call PressEnter()
Call PressEnter()
Call SetTextbox("Text","BSEG-SGTXT","",DT_AS01_0305_TEXT_OCC2,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS01_0305_PSTKY_OCC2,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0305_ACCOUNT_OCC2,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_AS01_0305_TTYPE_OCC2,False)
Call PressEnter()
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0305_AMOUNT_OCC3,False)
Call PressEnter()
Call PressEnter()
Call SetTextbox("Text","BSEG-SGTXT","",DT_AS01_0305_TEXT_OCC3,False)
Call ClickButton("Display Document Overview   \(Shift\+F2\)", False)
Call VerifyTextBoxContent("C", "RF05A-AZSAL", "", DT_AS01_0700_CHECK_TEXT_OF_C, False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)", False)
Call PressEnter()
Call TakeScreenShot
Call GetStatusBar("item1", "DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(Lcase(DT_AS01_0100_CHECK_TEXT_OF_STATUSBAR))
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



