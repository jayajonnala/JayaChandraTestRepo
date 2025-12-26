

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
'.................Test Script Name : Test_02GR02_001_Clear_LocVend_Int_Store_GR_DelivNot_cont_TradGoods
'.................Author : TCS 	   
'................ Creation Date 
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02GR02_001_Clear_LocVend_Int_Store_GR_DelivNot_cont_TradGoods"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR03_005_Loc_Vend_DSD_GR_Deliv_Note_w_Weighable_articles.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode MR11----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SelectCheckbox("PA_REUEB", 0, DT_MR11_1000_INVOICE_SURPLUS, False)
Call SetTextbox("Qty Var\. Less Than/Equal To","PA_DPROZ","","",False)   
Call SetTextbox("Purchase Order Date","RA_BEDAT-LOW","",ConvertDate(DT_MR11_1000_PURCHASE_ORDER_DATE),False)   
Call SetTextbox("to","RA_BEDAT-HIGH","",ConvertDate(DT_MR11_1000_TO),False) 
Call SetTextbox("Company Code","PA_BUKRS","",DT_MR11_1000_COMPANY_CODE,False)   
Call SetTextbox("Purchasing Document","RA_EBELN-LOW","",DT_MR11_1000_PURCHASING_DOCUMENT,False) 
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call ClickButton("Select All   \(F5\)",False)
Call TakeScreenShot()
Call ClickButton("Post Clearing   \(Ctrl\+F12\)",False)
Call TakeScreenShot()

Call GetGridContentByTitle("", 0, "Message Text", 1, "DT_MR11_500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_CMF_TEXT_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_MR11_500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_CMF_TEXT_OUTPUT",DT_MR11_500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_CMF_TEXT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyGridCellContent("", 1, "Message Text", 0, DT_MR11_500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_CMF_TEXT)


'----------------------Tcode we09----------------------------
Call SetTcode(DT_MR11_500_OKCD) 
Call PressEnter()  
Wait(2)

''Fill The Details and Execute
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call SetTextbox("Variant","V-LOW","","/ZMDIV_MIRO",True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SetTextbox("Logical Message","MESTYP-LOW","",DT_MR11_1000_LOGICAL_MESSAGE_OCC4,False)
Call SetTextbox("for Value \.\.\.","VALUE1_1","",DT_MR11_1000_FOR_VALUE_,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)

'''Navigate to Idoc Tree and Verify the details
Call GetLabelContentByRefLabel("IDoc number", 0, -48, "DT_IDOC_NUMBER_OUTPUT", False)
Call ClickLabel(DT_IDOC_NUMBER_OUTPUT,0,False)


Call ClickLinkGuiTree("IDoc "& DT_IDOC_NUMBER_OUTPUT&";Data records;ZZMDIV_E1RBKP","ZZMDIV_E1RBKP","",False)
Call VerifyTableCellContent(1,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_MR11_100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_0)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus()
