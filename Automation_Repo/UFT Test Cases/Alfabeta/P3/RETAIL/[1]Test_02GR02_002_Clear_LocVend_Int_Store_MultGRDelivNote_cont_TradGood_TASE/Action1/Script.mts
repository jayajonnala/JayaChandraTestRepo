
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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_02GR02_002_Clear_LocVend_Int_Store_MultGRDelivNote_cont_TradGood
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 24th Oct
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02GR02_002_Clear_LocVend_Int_Store_MultGRDelivNote_cont_TradGood"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR03_004_Loc_Vend_DSD_GR_Deliv_Note_w_Trading_Goods.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode MR11----------------------------

''Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Company Code","PA_BUKRS","",DT_MR11_1000_COMPANY_CODE,False)   
Call SetTextbox("Purchasing Document","RA_EBELN-LOW","",DT_TO,False)   
Call SetTextbox("to","RA_EBELN-HIGH","",DT_PURCHASING_DOCUMENT,False)   
Call SetTextbox("Purchase Order Date","RA_BEDAT-LOW","",DT_MR11_1000_PURCHASE_ORDER_DATE,False)   
Call SetTextbox("to","RA_BEDAT-HIGH","",DT_MR11_1000_TO_OCC2,False)
Call SelectCheckbox("PA_WEUEB",0,DT_MR11_1000_DELIVERY_SURPLUS,False)
Call SelectCheckbox("PA_REUEB",0,DT_MR11_1000_INVOICE_SURPLUS,False)
Call SelectRadioButton("PA_LISTE","Prepare List",False)

'Capture the screenshot
Call TakeScreenShot()

'Click on Execute Button
Call ClickButton("Execute   \(F8\)",False)
Wait(5)
'Capture the screenshot
Call TakeScreenShot()

'Select Al Records
Call ClickButton("Select All   \(F5\)",False)

Call ClickButton("Post Clearing   \(Ctrl\+F12\)",False)
Call TakeScreenShot()

'Get The Document No
Call GetGridContent("",0,"Message Text",1,"Application Area","M8","DT_MSG_TEXT_OUTPUT")

''----------------------Tcode WE09----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MR11_500_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Wait(2)
Call TakeScreenShot()

'Fill The Details and Execute
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","",DT_MR11_100_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SetTextboxNoLabel("VALUE1_1",0,DT_MR11_1000_FOR_VALUE,False)
Call SetTextbox("Logical Message","MESTYP-LOW","",DT_MR11_1000_LOGICAL_MESSAGE_OCC4,False)
Wait(2)
Call TakeScreenShot()

'Verify Text Box Content
Call VerifyTextBoxContent("Logical Message","MESTYP-LOW",0,DT_MR11_1000_LOGICAL_MESSAGE_OCC4,False)

Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call ClickButtonIfExist("Yes",True)
Wait(5)
Call VerifyStatusBarMessageType("S")
Call TakeScreenShot()
Call VerifyStatusBar("IDocs were found")

'Verify If GUI Label Exists
Call VerifyifGuiLabelExists(DT_MR11_120_CHECK_TEXT_OF_DATA_PASSED_TO_PORT_OK)

'Navigate to Idoc Tree and Verify the details
Call GetLabelContentByRefLabel(Right(DT_MR11_1000_CHECK_TEXT_OF_LOGICAL_MESSAGE,Len(DT_MR11_1000_CHECK_TEXT_OF_LOGICAL_MESSAGE)),994,0,"DT_IDOC_NUMBER",False)
Call ClickLabel(DT_IDOC_NUMBER,0,False)

Call ClickLinkGuiTree("IDoc "& DT_IDOC_NUMBER&";Data records;ZZMDIV_E1RBKP","ZZMDIV_E1RBKP","",False)
Call TakeScreenShot()
Call ClickLinkGuiTree("IDoc "& DT_IDOC_NUMBER&";Data records;ZZMDIV_E1RSEG","ZZMDIV_E1RSEG","",False)
Call VerifyTableCellContent(4,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_FLD_CONT_3_OCC2)''DT_MR11_100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_3)
Call ClickLinkGuiTree("IDoc "& DT_IDOC_NUMBER&";Data records;ZZMDIV_E1RSEG","ZZMDIV_E1RSEG","",False)
Call VerifyTableCellContent(4,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_FLD_CONT_3_OCC2)''DT_MR11_100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_3_OCC2)

'Exit the TRansaction
'Call SetTcode(DT_MR11_100_OKCD) 
'Call PressEnter() 
Call ClickButton("Exit   (Shift+F3)",False)
Call ClickButton("Exit   (Shift+F3)",False)
Call ClickButton("Exit   (Shift+F3)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
'Wait(2)
'Call VerifyStatusBarMessageType("E")

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

