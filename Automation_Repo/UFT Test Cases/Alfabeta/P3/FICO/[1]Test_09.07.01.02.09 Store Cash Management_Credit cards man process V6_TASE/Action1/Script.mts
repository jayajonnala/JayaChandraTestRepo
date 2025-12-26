

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.02.09 Store Cash Management_Credit cards man process V6_TASE
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_09.07.01.02.09 Store Cash Management_Credit cards man process V6_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'


'----------------------Tcode FB50----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FB50_1000_COMPANY_CODE)
'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB50_1000_COMPANY_CODE,True)
Call TakeScreenShot() 
Call ClickButtonIfExist("Continue   \(Enter\)",true)
'Call ClickButton("Continue   \(Enter\)",True)
Wait(2)
Call SetTextbox("Doc\.Header Text","ACGL_HEAD-BKTXT","",DT_FB50_1010_DOCHEADER_TEXT,False)
Call SetTextbox("Reference","ACGL_HEAD-XBLNR","",DT_FB50_1010_REFERENCE,False)
Call SetTextbox("Document Date","ACGL_HEAD-BLDAT","",DT_FB50_1010_DOCUMENT_DATE,False)
Call TakeScreenShot()

Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",1,DT_FB50_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",2,DT_FB50_0100_TABLECELL_GL_ACCT_1,False)

Call SelectCellGuiTable("SAPLFSKBTABLE", "D/C", "G/L acct", DT_FB50_0100_TABLECELL_GL_ACCT_0, False)
Call SendKey("{F4}")
Call SendKey("{TAB}")

Call SelectCellGuiTable("SAPLFSKBTABLE", "D/C", "G/L acct", DT_FB50_0100_TABLECELL_GL_ACCT_1, False)
Call SendKey("{F4}")
Wait 2
Call SendKey("{DOWN}")
Call SendKey("{TAB}")

Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",1,DT_FB50_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",2,DT_FB50_0100_TABLECELL_AMOUNT_IN_DOCCURR_1,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Tax code",2,DT_FB50_0100_TABLECELL_TAX_CODE_1,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Tax jurisdictn code",2,DT_FB50_0100_TABLECELL_TAX_JURISDICTN_CODE_1,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Assignment",1,DT_FB50_0100_TABLECELL_ASSIGNMENT_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Text",1,DT_FB50_0100_TABLECELL_TEXT_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Text",2,DT_FB50_0100_TABLECELL_TEXT_1,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",1,DT_FB50_0100_TABLECELL_BUSINESS_AREA_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",2,DT_FB50_0100_TABLECELL_BUSINESS_AREA_1,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Cost center",2,DT_FB50_0100_TABLECELL_COST_CENTER_1,False)

Call ClickButtonIfExist("Simulate Document Posting   \(F9\)",fALSE)
Call ClickButton("Post   \(Ctrl\+S\)",fALSE)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",fAlse)
wait 3 
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call GetStatusBar("item1","DT_DOC_NR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NR_OUTPUT",DT_DOC_NR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickButtonifExist("Cancel   \(F12\)",True)
Call ClickButtonifExist("Yes",True)
'
'----------------------Tcode FAGLL03----------------------------


Call SetTcode(DT_FB50_0100_OKCD)
Call PressEnter() 

Call SelectRadioButton("X_AISEL","All Items",False)
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDate(DT_FB50_1000_POSTING_DATE),False)
Call SetTextbox("to","SO_BUDAT-HIGH","",ConvertDate(DT_FB50_1000_TO),False)
Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FB50_1000_GL_ACCOUNT,False)
Call TakeScreenShot()

Call FocusTextBox("to","SD_SAKNR-HIGH",False)
Call ClickButton("Execute   \(F8\)",fALSE)
Call  SelectColumnGuiGrid("", 1, "Document Number", False)
''''Call ClickLabel("DocumentNo",1,False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Wait(2)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FB50_1105_DOCUMENT_NUMBER,True)
Call TakeScreenShot()

Call ClickButton("Execute   \(Enter\)",True)
wait(2)
Call VerifyGridCellContent("", 1, "BELNR", "", DT_FB50_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BLART", "", DT_FB50_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "MWSKZ", "", DT_FB50_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("", 1, "PRCTR", "", DT_FB50_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)


'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB50_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR,"wnd\[0\]/usr/lbl\[9,8\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB50_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART,"wnd\[0\]/usr/lbl\[25,8\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB50_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ,"wnd\[0\]/usr/lbl\[70,8\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB50_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR,"wnd\[0\]/usr/lbl\[84,8\]")

'''------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()



