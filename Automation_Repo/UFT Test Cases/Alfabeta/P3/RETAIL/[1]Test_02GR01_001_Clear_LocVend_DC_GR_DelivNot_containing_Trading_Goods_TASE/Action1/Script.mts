

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02GR01_001_Clear_LocVend_DC_GR_DelivNot_containing_Trading_Goods
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

gstrTestCaseName = "Test_02GR01_001_Clear_LocVend_DC_GR_DelivNot_containing_Trading_Goods"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''--------------------------------MR11-----------------------------


Call SetTextbox("Company Code","PA_BUKRS","",DT_MR11_1000_COMPANY_CODE,False)
Call SetTextbox("Purchasing Document","RA_EBELN-LOW","",DT_MR11_1000_PURCHASING_DOCUMENT,False)
Call SelectCheckbox("PA_REUEB",0,DT_MR11_1000_INVOICE_SURPLUS,False)
Call SetTextbox("Purchase Order Date","RA_BEDAT-LOW","",ConvertDate(DT_MR11_1000_PURCHASE_ORDER_DATE),False)
Call SetTextbox("to","RA_BEDAT-HIGH","",ConvertDate(DT_MR11_1000_TO),False)
Call SetTextbox("Item","RA_EBELP-LOW","",FormatBlank(DT_MR11_1000_ITEM),False)
Call SetTextbox("Qty Var\. Less Than/Equal To","PA_DPROZ","",FormatBlank(DT_MR11_1000_QTY_VAR_LESS_THANEQUAL_TO),False)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Call ClickButton("Select All   \(F5\)",False)
Call ClickButton("Post Clearing   \(Ctrl\+F12\)",False)
Call TakeScreenShot()
Call GetGridContent("",0,"Message Text",1,"<NA>","<NA>","DT_MR11_500_GETCELLVALUE_OF_GRIDCELL_0_CMF_TEXT_OUTPUT")

''--------------------------------WE09-----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MR11_500_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","",FormatBlank(DT_MR11_100_VARIANT),True)
Call SetTextbox("Created By","ENAME-LOW","",FormatBlank(DT_MR11_100_CREATED_BY_OCC2),True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 

Call SetTextbox("Logical Message","MESTYP-LOW","",DT_MR11_1000_LOGICAL_MESSAGE_OCC4,False)
Call SetTextbox("for Value \.\.\.","VALUE1_1","",DT_MR11_1000_FOR_VALUE,False)
Call VerifyTextBoxContent("Logical Message","MESTYP-LOW",0,DT_MR11_1000_CHECK_TEXT_OF_LOGICAL_MESSAGE,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Call VerifyStatusBarMessageType("S")

Call VerifyStatusBar("IDocs were found")
Call SetFocusGuiLabel("",32,72,False)
Call SendKey("{F2}")
Wait 2
Call ClickLinkGuiTree( "#1;#2;#1","#1",0,False)
Call VerifyTableCellContent(1,"Fld cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_MR11_100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_0)
Call VerifyTableCellContentByRef("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Name","BELNR","Fld cont.",DT_MR11_100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_0)

Call LogOff()
Call FinalStatus ()


