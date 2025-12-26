
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : [1]Test_Replenishment with Auction PO_P5 Reprocess IDoc_TASE
'.................Test Scenario: AT_P2P_Fresh_Replenishment via Workbench and Auction PO
'.................TCode: BD87, MB51
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

gstrTestCaseName = "[1]Test_Replenishment with Auction PO_P5 Reprocess IDoc_TASE"
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

'----------------------Tcode BD87----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

' SetTextbox(textboxAttachedText, textboxName, textboxIndex, textboxValue, blnIsItPopup)
Call SetTextbox("IDoc Number","SX_DOCNU-LOW",0,DT_BD87_1100_IDOC_NUMBER,False)
Call SetTextbox("Changed On","SX_UPDDA-LOW",0,ConvertDate(DT_BD87_1100_CHANGED_ON),False)
Call TakeScreenShot()

' ClickButton(tooltipOrButtonName, blnIsItPopup)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

' SelectNodeGuiTree(treeIndex, itemPath)
Call SelectNodeGuiTree(0,"#2;#1;#1")

' SelectItemGuiTree(treeIndex, itemPath, itemText)
'Call SelectItemGuiTree(0,"#2;#1;#1","Application document not posted")
Call ActivateNodeGuiTree(0,"#2;#1;#1")


'Call ClickButton("Process Selected Node   \(F8\)",False)
'Call TakeScreenShot()

' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContent("IDoc Selection",1,"IDoc status",0,DT_BD87_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NEW_STAT)

Call VerifyGridCellContent("IDoc Selection",1,"Status Text",0,Lcase(DT_BD87_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_STAT_TXT))


'----------------------Tcode MB51----------------------------

Call SetTcode(DT_BD87_0100_OKCD) 
Call PressEnter()    
Call TakeScreenShot()

Call SetTextbox("Purchase order","EBELN-LOW",0,DT_BD87_1000_PURCHASE_ORDER,False)
Call SetTextbox("Site","WERKS-LOW",0,DT_BD87_1000_SITE,False)
Call SetTextbox("Order","AUFNR-LOW",0,"",False)
Call SetTextbox("Posting Date","BUDAT-LOW",0,ConvertDate(DT_BD87_1000_POSTING_DATE),False)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContent("",1,"Movement type",0,DT_BD87_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWART)
Call VerifyGridCellContent("",1,"Reference",0,DT_CHECK_PO)
Call VerifyGridCellContent("",1,"Qty in unit of entry",0,DT_BD87_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ERFMG)

'Call VerifyifGuiLabelExistsByRelativeid(DT_BD87_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWART,"wnd\[0\]/usr/lbl\[6,5\]")
'Call SetHorizontalScrollBar(120,False)
'Wait(10)
'Call VerifyifGuiLabelExistsByRelativeid(DT_BD87_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ERFMG,"wnd\[0\]/usr/lbl\[161,5\]")
'Call VerifyifGuiLabelExistsByRelativeid(DT_CHECK_PO,"wnd\[0\]/usr/lbl\[115,5\]")


Call LogOff()
Call FinalStatus()

'*********************************************End Of Script*********************************************************************

