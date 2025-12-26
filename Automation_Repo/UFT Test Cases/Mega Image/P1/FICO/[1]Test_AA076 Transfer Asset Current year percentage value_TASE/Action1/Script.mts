
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA076 Transfer Asset Current year percentage value
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


gstrTestCaseName = "Test_AA076 Transfer Asset Current year percentage value"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA010- Asset Creation - create asset directly in SAP.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''--------TransactionCode-abumn----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot
Call SetTextboxPopupIfExist("SVALD-VALUE","Company Code","RO02")
'Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code","RO02")
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("Asset","RAIFP2-ANLN1","",DT_ABUMN_0300_ASSET,False)
Call SetTextbox("Asset","RAIFP2-ANLN2","",DT_ABUMN_0320_EXISTING_ASSET_OCC1,False)
Call TakeScreenShot

Call SetTextbox("Text","RAIFP2-SGTXT","",DT_ABUMN_0206_TEXT,False)
Call TakeScreenShot

Call SetTextbox("Existing asset","RAIFP3-ANLN1","",DT_ABUMN_0320_EXISTING_ASSET,False)
Call SetTextbox("Existing asset","RAIFP3-ANLN2","",DT_AS01_0320_EXISTING_ASSET_OCC1,False)
Call TakeScreenShot
''Call SetTextbox("Asset Value Date","RAIFP1-BZDAT","",ConvertDate(DT_ABUMN_0202_ASSET_VALUE_DATE),False)
''Call TakeScreenShot
''Call SetTextbox("Posting Date","RAIFP1-BUDAT","",ConvertDate(DT_ABUMN_0201_POSTING_DATE),False)
''Call TakeScreenShot
''Call SetTextbox("Document Date","RAIFP1-BLDAT","",ConvertDate(DT_ABUMN_0200_DOCUMENT_DATE),False)
''Call TakeScreenShot
Call SetTextbox("Asset Value Date","RAIFP1-BZDAT","",DT_ABUMN_0202_ASSET_VALUE_DATE,False)
Call TakeScreenShot
''Call SetTextbox("Posting Date","RAIFP1-BUDAT","",DT_ABUMN_0201_POSTING_DATE,False)
Call SetTextbox("Posting Date","RAIFP1-BUDAT","",DT_ABUMN_0202_ASSET_VALUE_DATE,False)
Call TakeScreenShot
''Call SetTextbox("Document Date","RAIFP1-BLDAT","",DT_ABUMN_0200_DOCUMENT_DATE,False)
Call SetTextbox("Document Date","RAIFP1-BLDAT","",DT_ABUMN_0202_ASSET_VALUE_DATE,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Additional Details", False)
Call TakeScreenShot
Call SetTextbox("Document type","RAIFP1-BLART","",DT_ABUMN_0204_DOCUMENT_TYPE,False)
Call TakeScreenShot
Call SetTextbox("Posting period","RAIFP2-MONAT","",DT_ABUMN_0203_POSTING_PERIOD,False)
Call TakeScreenShot
Call SetTextbox("Transfer variant","RAIFP1-TRAVA","",DT_ABUMN_0210_TRANSFER_VARIANT,False)
Call TakeScreenShot
Call SetTextbox("Reference","RAIFP1-XBLNR","",DT_ABUMN_0207_REFERENCE,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call SelectTab("TABSTRIP100", "Partial Transfer", False)
Call TakeScreenShot
Call SelectRadioButton("RAIFP2-XANEU","From curr\.-yr aquis\.",False)
Call TakeScreenShot
Call SetTextbox("Percentage rate","RAIFP2-PROZS","",DT_ABUMN_0401_PERCENTAGE_RATE,False)
Call TakeScreenShot
Call PressEnter()   
Call TakeScreenShot

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item2", "DT_ABUMN_0100_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Asset transaction posted with document no. "& DT_ABUMN_0300_COMPANY_CODE &" "&DT_ABUMN_0100_CHECK_TEXT_OF_STATUSBAR_OUTPUT)


''''''--------TransactionCode-AS03----------''''
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_ABUMN_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Asset","ANLA-ANLN1","",DT_ABUMN_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_ABUMN_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_ABUMN_0100_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0,FormatLocalDecimal( DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC1))

''------------------------------------------------AS03----------------------------------------
Call SetTcode(DT_ABUMN_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Asset","ANLA-ANLN1","",DT_ABUMN_0100_ASSET_OCC1,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_ABUMN_0100_SUBNUMBER_OCC1,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_ABUMN_0100_COMPANY_CODE_OCC1,False)
Call TakeScreenShot
Call PressEnter()  
Call TakeScreenShot
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0,FormatLocalDecimal( DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC1))


Call LogOff()
Call FinalStatus ()

Public Function FormatLocalDecimal(Dtnumber)
On error Resume Next
			Dtnumber=Cstr(Dtnumber)
			FormatLocalDecimal =Left(Dtnumber,Len(Dtnumber)-3)& "."&Right(Dtnumber,3) &",00"	
'		

End Function



