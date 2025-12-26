'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Correspondence with Customer_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 08th March
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

gstrTestCaseName = "Test_Correspondence with Customer_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Correspondence with Customer_p2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode FB03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB03_0100_FISCAL_YEAR,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB03_0100_COMPANY_CODE,False)
Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB03_0100_DOCUMENT_NUMBER,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()
'verify doc nummber
Call VerifyTextBoxContent("Document Number","BKPF-BELNR",0,DT_FB03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER,False)
'verify the grid coponents
call VerifyGridCellContent("",1,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",1,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",2,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
call VerifyGridCellContent("",2,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
call VerifyGridCellContent("",3,"Posting Key","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
call VerifyGridCellContent("",3,"Account","",DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
'Capture the screenshot
Call TakeScreenShot()
'
'''----------------------Tcode ZFIAR_RS_F17P----------------------------
'
'Enter the Tcode
Call SetTcode(DT_FB03_0750_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_FB03_0750_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer","DD_KUNNR-LOW","",DT_FB03_1000_CUSTOMER,False)

Public Function VariantPrint_bit_1()

'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Choose   \(F2\)",True)
Wait(1)
Call SetTextbox("Customer","DD_KUNNR-LOW","",DT_FB03_1000_CUSTOMER,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
'wait and Capture the screenshot
Wait(2)
Call TakeScreenShot()

End Function

Public Function VariantPrint_bit()

Call VariantPrint_bit_1()
Call ClickButton("Cancel   \(F12\)",False)
'wait and Capture the screenshot
Wait(2)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
'wait and Capture the screenshot
Wait(2)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
'wait and Capture the screenshot
Wait(2)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
'wait and Capture the screenshot
Wait(1)
Call TakeScreenShot()

End Function
'--------------------------------------------------Variant 1---------------------------------------------
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
'select required variant and choose
Call SelectRowGuiGrid("Variant Catalog for Program ZFIAR_RS_SAPF130D_PDF (7)","",DT_REF_COLUMN,DT_REF_VALUE,True)
Call VariantPrint_bit()
'--------------------------------------------------Variant 2---------------------------------------------
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
'select required variant and choose
Call SelectRowGuiGrid("Variant Catalog for Program ZFIAR_RS_SAPF130D_PDF (7)","",DT_REF_COLUMN,DT_REF_VALUE_2,True)
Call VariantPrint_bit()
'--------------------------------------------------Variant 3---------------------------------------------
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
'select required variant and choose
'Call SelectRowGuiGrid("Variant Catalog for Program ZFIAR_RS_SAPF130D_PDF","",DT_REF_COLUMN,DT_REF_VALUE_3,True)
Call SelectRowGuiGrid("Variant Catalog for Program ZFIAR_RS_SAPF130D_PDF (7)","",DT_REF_COLUMN,DT_REF_VALUE_3,True)
Call VariantPrint_bit()
'--------------------------------------------------Variant 4---------------------------------------------
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
'select required variant and choose
Call SelectRowGuiGrid("Variant Catalog for Program ZFIAR_RS_SAPF130D_PDF (7)","",DT_REF_COLUMN,DT_REF_VALUE_4,True)
Call VariantPrint_bit_1()

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


