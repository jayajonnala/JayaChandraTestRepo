
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.1.1.13.5. Maintain Bill of Materialof Article - Class 1, Class
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.1.1.13.5. Maintain Bill of Materialof Article - Class 1, Class"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.6.1.7.1. Check Price Change File from Workflow.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode CS01 ----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("BOM Usage","RC29N-STLAN","",DT_CS01_0100_BOM_USAGE,False)
Call SetTextbox("Valid From","RC29N-DATUV","",ConvertDate(DT_CS01_0100_VALID_FROM),False)
Call SetTextbox("Article","RC29N-MATNR","",DT_CS01_0100_ARTICLE,False)
Call FocusTextBox("Change Number","RC29N-AENNR", False)
Call TakeScreenShot()

Call PressEnter()  
Call TakeScreenShot()

Call PressEnter()  
Call TakeScreenShot()

Call PressEnter()  
Call TakeScreenShot()

Call PressEnter()  
Call TakeScreenShot()


Call SetTableData("SAPLCSDITCMAT", "ICt", 1, "", "", DT_CS01_0152_TABLECELL_ICT_0, False)
Call SetTableData("SAPLCSDITCMAT", "ICt", 2, "", "", DT_CS01_0152_TABLECELL_ICT_1, False)

Call SetTableData("SAPLCSDITCMAT", "Component", 1, "", "", DT_CS01_0152_TABLECELL_COMPONENT_0, False)
Call SetTableData("SAPLCSDITCMAT", "Component", 2, "", "", DT_CS01_0152_TABLECELL_COMPONENT_1, False)

Call SetTableData("SAPLCSDITCMAT", "Quantity", 1, "", "", DT_CS01_0152_TABLECELL_QUANTITY_0, False)
Call SetTableData("SAPLCSDITCMAT", "Quantity", 2, "", "", DT_CS01_0152_TABLECELL_QUANTITY_1, False)

Call SetTableData("SAPLCSDITCMAT", "Un", 1, "", "", DT_CS01_0152_TABLECELL_UN_0, False)
Call SetTableData("SAPLCSDITCMAT", "Un", 2, "", "", DT_CS01_0152_TABLECELL_UN_1, False)

Call SelectRowGuiTable("SAPLCSDITCMAT", "Component", DT_CS01_0152_TABLECELL_COMPONENT_0, False)
Call SelectRowGuiTable("SAPLCSDITCMAT", "Component", DT_CS01_0152_TABLECELL_COMPONENT_1, False)

Call ClickButton("Item   \(F7\)",False)

Call SelectCheckbox("RC29P-SANKO", 0, DT_CS01_0840_ENGINEERINGDESIGN, False)
Call PressEnter()
Call TakeScreenShot()
Call SelectCheckbox("RC29P-SANKO", 0, DT_CS01_0840_ENGINEERINGDESIGN_OCC1, False)
Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Header   \(F6\)",False)
Call SelectTab("TS_HEAD", "Addnl Data", False)
Call TakeScreenShot()
Call SetTextbox("Lab/Office","RC29K-LABOR","",DT_CS01_1111_LABOFFICE,False)
Call PressEnter()
Call TakeScreenShot()

Call VerifyTextBoxContent("Lab/Office","RC29K-LBTXT", 0,Lcase(DT_CS01_1111_CHECK_TEXT_OF_LABOFFICE),False)

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBar(DT_CS01_0100_CHECK_TEXT_OF_STATUSBAR)

Call LogOff'
Call FinalStatus ()
