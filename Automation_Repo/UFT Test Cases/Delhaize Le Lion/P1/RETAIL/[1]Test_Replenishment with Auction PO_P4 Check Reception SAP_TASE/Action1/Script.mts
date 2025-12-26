
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : [1]Test_Replenishment with Auction PO_P4 Check Reception SAP_TASE
'.................Test Scenario: AT_P2P_Fresh_Replenishment via Workbench and Auction PO
'.................TCode: ME23N, MEK1
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

gstrTestCaseName = "Test_Replenishment with Auction PO_P4 Check Reception SAP_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'ensure no open session
''''Call CloseSessionsSAP()
'Login to SAP System

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode WE02----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextbox("Message Variant","MESCOD-LOW",0,DT_WE02_1100_MESSAGE_VARIANT,False)
Call SetTextbox("Message Function","MESFCT-LOW",0,DT_WE02_1100_MESSAGE_FUNCTION,False)
Call SetTextbox("Partner Number","PPPRN-LOW",0,DT_WE02_1100_PARTNER_NUMBER,False)

Call SelectTab("TABSTRIP_IDOCTABBL","EDI",False)
Call TakeScreenShot()
Call SetTextbox("Reference to Message Group","REFGRP-LOW",0,DT_WE02_1300_REFERENCE_TO_MESSAGE_GROUP,False)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot()

'Call VerifyTextBoxContent("Current Status","EDIDC-STATUS",0,DT_WE02_0100_CHECK_TEXT_OF_CURRENT_STATUS,False)
Call TakeScreenShot()
Call GetWindowValue("DT_GET_IDOC_NUMBER_OUTPUT", False)
Call TakeScreenShot()

'Call GetNodeTextGuiTree(0,"idoc",0,"DT_GET_IDOC_NUMBER_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_GET_IDOC_NUMBER_OUTPUT",DT_GET_IDOC_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ActivateNodeGuiTree(0,"#1;#3;#1")
Call Takescreenshot()
 ''VerifyTextBoxContent(textboxAttachedText, textboxName, textboxIndex, expectedValue, blnIsItPopup)
''Call VerifyTextBoxContent("Message","T100-TEXT",0,DT_WE02_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1,False)
''Call VerifyNodeTextGuiTree(0,DT_WE02_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1)
 ''VerifyNodeTextGuiTree(treeIndex, itemText)

''----------------------Tcode SE38----------------------------

Call SetTcode(DT_WE02_0100_OKCD) 
Call PressEnter()
Call TakeScreenShot()

Call SetTextbox("Program","RS38M-PROGRAMM",0,DT_WE02_0100_PROGRAM,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot()

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)", False)

Call SelectRowGuiGrid("Variant Catalog for Program RMEBEIN1",0,"Variant name","DLL_AUCTION",True)
Call ClickButton("Choose   \(F2\)", True)


Call SetTextbox("Purchasing Documents","S_EBELN-LOW",0,DT_WE02_1000_PURCHASING_DOCUMENTS,False)
Call SetTextbox("Vendor","S_LIFNR-LOW",0,DT_WE02_1000_VENDOR,False)
Call SetTextbox("Material","S_MATNR-LOW",0,DT_WE02_1000_ARTICLE,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot()

Call VerifyTextBoxContent("Information Message","MESSTXT1",0,Lcase(DT_WE02_0010_CHECK_TEXT_OF_MESSTXT1),True)
Call ClickButton("Continue   \(Enter\)", False)
Call TakeScreenShot()
Call ClickButton("Yes", False)
Call TakeScreenShot()
Call ClickButtonifexist("Continue   \(Enter\)", False)
''' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContent("",1,"Message Text",0,Lcase(DT_WE02_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_0_T_MSG))


'----------------------Tcode SE38----------------------------

Call SetTcode(DT_WE02_0100_OKCD_OCC1) 
Call PressEnter()
Call TakeScreenShot()

Call SetTextbox("Program","RS38M-PROGRAMM",0,DT_WE02_0100_PROGRAM_OCC1,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot()
Call ClickButton("Yes", False)
Call TakeScreenShot()



'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

