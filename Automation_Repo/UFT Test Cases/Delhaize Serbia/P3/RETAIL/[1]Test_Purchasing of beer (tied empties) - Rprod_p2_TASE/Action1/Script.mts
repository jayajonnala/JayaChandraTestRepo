'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : DT_DSD Purchasing of beer (tied empties) - Rprod_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 25th Feb
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

gstrTestCaseName = "Test_DSD Purchasing of beer (tied empties) - Rprod_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_DSD Purchasing of beer (tied empties) - Rprod_p2.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''----------------------Tcode MIGO----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()
'Enter the article doc and Press Enter
Call SetCombo("GODYNPRO-ACTION","Display")
'Call SetCombo("GODYNPRO-REFDOC","Article Document")
Call SetComboByKey("GODYNPRO-REFDOC","R02")
'Capture the screenshot
Call TakeScreenShot()
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_MIGO_2010_GODYNPROMAT_DOC,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()
Call SelectTab("TS_GOHEAD","Doc. info",False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("FI Documents",False)
'Capture the screenshot
Call TakeScreenShot()
Call GetCellDataGuiGridPopupByRefTwoColumns("Documents in Accounting","","Document Number","Object type text","Accounting document","Object type text","Accounting document","DT_MIGO_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SelectCellGuiGrid("Documents in Accounting","",DT_MIGO_0200_GRIDCELL_0_DOC_NUMBER,"Document Number",true)
Call SelectRowGuiGridbyRowNo("Documents in Accounting","",DT_MIGO_0200_GRIDCELL_0_DOC_NUMBER,true)
''Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Display Document   \(F2\)",True)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()
'''-------------------------------------------------MIRO-----------------------------------------
Call SetTcode(DT_MIGO_0750_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MIGO_0750_OKCD)
Call SetCombo("RM08M-VORGANG","Invoice")
Wait(2)
'
Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_MIGO_0010_INVOICE_DATE),False)
wait(1)
Call PressEnter() 
wait(1)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIGO_0010_REFERENCE1,False)
Call PressEnter()
Call SetTextbox("Incg Doc. Nmbr","INVFO-INWARDNO_HD","",DT_MIGO_2010_GODYNPROMAT_DOC,False)
Call PressEnter()
'Enter the delivery Note No
Call SetCombo("RM08M-REFERENZBELEGTYP",DT_MIGO_6020_RM08MREFERENZBELEGTYP)
'Select Calculate Tax field as true
Call SelectCheckbox("INVFO-XMWST",0,DT_MIGO_0010_CALCULATE_TAX,False)
Call TakeScreenShot()
Call SetTextboxNoLabel("RM08M-LFSNR",0,DT_MIGO_6212_RM08MLFSNR1,False)
Call PressEnter()
Call TakeScreenShot()
Call SetTableData("SAPLMR1MTC_MR1M","Amount",2,"Item",2,DT_MIGO_6310_TABLECELL_AMOUNT_1,False)
Call PressEnter()
Call TakeScreenShot()
'Get the remaining balance and enter it in Amount Field
'Call GetTextboxValue("RM08M-DIFFERENZ",0,"DT_MIGO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT",False)
Call GetTextboxValue("RM08M-DIFFERENZ",1,"DT_MIGO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIGO_0010_AMOUNT,False)
Call PressEnter()
Wait(1)
Call TakeScreenShot()
'Click on Post Buton
Call SelectMenuBar("Invoice Document;Post")
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If invoice is generated
Call GetStatusBar("item1","DT_MIGO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_MIGO_6000_CHECK_TEXT_OF_STATUSBAR)
'-----------------------------------------------MIR4-----------------------------------------------------
Call SetTcode(DT_MIGO_6000_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_MIGO_6000_OKCD)
Call TakeScreenShot()
'Display the Invoice Details
Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MIGO_6150_INVOICE_DOCUMENT_NO,False)
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",DT_MIGO_6150_FISCAL_YEAR,False)
Call TakeScreenShot()
Call PressEnter()
wait(1)
Call TakeScreenShot()
'Click on Follow On Document
Call ClickButtonIfExist("Follow-On Documents \.\.\.   \(F8\)",False)
wait(1)
Call TakeScreenShot()
Call SelectCellGuiGrid("Documents in Accounting","",DT_MIGO_0200_GRIDCELL_0_DOC_NUMBER_OCC1,"Document Number",true)
Call SelectRowGuiGridbyRowNo("Documents in Accounting","",DT_MIGO_0200_GRIDCELL_0_DOC_NUMBER_OCC1,true)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Display Document   \(F2\)",True)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Exit   \(Shift\+F3\)",False)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Exit   \(Shift\+F3\)",False)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

