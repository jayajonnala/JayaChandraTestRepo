
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04-01-05-01- Automatic PO for Advanced Stock_Mult_Art_P3
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

gstrTestCaseName = "Test_04-01-05-01- Automatic PO for Advanced Stock_Mult_Art_P3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\DLL\RETAIL\DT_04-03-03-01 GDSN article with Fost fresh_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario



'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

''--------TransactionCode-WE02----------''''
' 
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call SetTextbox("Changed At","SX_UPDTI-LOW","",DT_BD87_1100_CHANGED_AT,False)
Call SetTextbox("to","SX_UPDTI-HIGH","",DT_BD87_1100_TO_OCC1,False)
Call SetTextbox("Changed On","SX_UPDDA-LOW","",ConvertDate(DT_BD87_1100_CHANGED_ON),False)
Call SetTextbox("to","SX_UPDDA-HIGH","",ConvertDate(DT_BD87_1100_TO),False)
Call SetTextbox("Message Type","SX_MESTY-LOW","",DT_BD87_1100_MESSAGE_TYPE,False)
Call SetTextbox("IDoc Status","SX_STATU-LOW","",DT_BD87_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_STATUS,False)
Call SetTextbox("Partner System","SX_PRPRN-LOW","",DT_BD87_1100_PARTNER_SYSTEM,False)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)

Call TakeScreenShot()
Call ActivateItemGuiTree(0,"#2;#1;#1","#1")
Call VerifyGridCellContent("IDoc Selection",1,"STATUS",0,DT_BD87_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_STATUS)
Call VerifyGridCellContent("IDoc Selection",1,"MESTYP",0,DT_BD87_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MESTYP)
Call VerifyGridCellContent("IDoc Selection",1,"MESCOD",0,DT_BD87_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MESCOD)
Call VerifyGridCellContent("IDoc Selection",1,"MESFCT",0,DT_BD87_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MESFCT)
Call SelectRowGuiGridbyRowNo("IDoc Selection",0,1,False)
Call ClickButton("Display Related Objects   \(Shift\+F8\)",False)
Call TakeScreenShot()
Call SelectCellGuiTable("SAPLSRELNEIGHBOR_LIST","Object ID","Obj. type","Goods Movement",False)
Call ClickButton("Choose   \(F2\)",True)
Call SelectTab("TS_GOITEM", "Quantity", False)
Call TakeScreenShot()
Call SelectTab("TS_GOITEM", "Where", False)
Call VerifyTextBoxContent("Movement type", "GOITEM-BWART", "",DT_BD87_0325_CHECK_TEXT_OF_MOVEMENT_TYPE, False)
Call SelectTab("TS_GOITEM", "Purchase Order Data", False)
Call GetTextboxValue("GODYNPRO-MAT_DOC",0,"DT_BD87_2010_CHECK_TEXT_OF_GODYNPROMAT_DOC_OUTPUT",False)
Call GetTextboxValue("GOITEM-EBELN",0,"DT_PO_NUMBER_OUTPUT",False)
Call TakeScreenShot()


''--------TransactionCode-BD87----------''''
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet) 
Call SetTcode(DT_ME23N_TRANSACTION)     
Call PressEnter()     
Call CheckTCodeScreen(DT_ME23N_TRANSACTION)

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call TakeScreenShot
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_PO_NUMBER_OUTPUT,True) 
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur\. Order",True) 
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot

Call VerifyTextBoxNoLabelContent("MEPO_TOPLINE-EBELN",0,DT_PO_NUMBER_OUTPUT,False)
Call VerifyTextBoxNoLabelContent("MEPO_TOPLINE-SUPERFIELD",0,DT_VENDOR,False)

Call LogOff()
Call FinalStatus ()

