'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_Customer : Test_PRE1_ME21N_Create_PO_TASE
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE1_ME21N_Create_PO_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DLL\FICO\TASE_DT_PRE1_ME21N_Create_PO.xls"

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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)



'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'--------TransactionCode-ZFIAR_AFF_CLEARING----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()

Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD",0,DT_ME21N_1105_VENDOR,False) 

Call PressEnter()
Call SelectTab("HEADER_DETAIL","Org. Data",False)

Call SetTextbox("Purch. Org.","MEPO1222-EKORG",0,DT_ME21N_1221_PURCH_ORG,False)     
Call SetTextbox("Purch. Group","MEPO1222-EKGRP",0,DT_ME21N_1221_PURCH_GROUP,False)     
Call SetTextbox("Company Code","MEPO1222-BUKRS",0,DT_ME21N_1221_COMPANY_CODE,False)    
   
'''Call PressEnter()     ' - Line (7)

Call SetTableData("SAPLMEGUITC_1211","Itm","1","","",DT_ME21N_1211_TABLECELL_ITM_0,False)     
Call SetTableData("SAPLMEGUITC_1211","A","1","","",DT_ME21N_1211_TABLECELL_A_0,False)     
Call SetTableData("SAPLMEGUITC_1211","Article","1","","","",False)     
Call SetTableData("SAPLMEGUITC_1211","Short Text","1","","",DT_ME21N_1211_TABLECELL_SHORT_TEXT_0,False)     
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)    
Call SetTableData("SAPLMEGUITC_1211","OUn","1","","",DT_ME21N_1211_TABLECELL_OUN_0,False)     
Call SetTableData("SAPLMEGUITC_1211","C","1","","",DT_ME21N_1211_TABLECELL_C_0,False)     
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",DT_ME21N_1211_TABLECELL_DELIV_DATE_0,False)     
Call SetTableData("SAPLMEGUITC_1211","Net Price","1","","",DT_ME21N_1211_TABLECELL_NET_PRICE_0,False)  
Call SetTableData("SAPLMEGUITC_1211","Currency","1","","",DT_ME21N_1211_TABLECELL_CURRENCY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Per","1","","",DT_ME21N_1211_TABLECELL_PER_0,False)     
Call SetTableData("SAPLMEGUITC_1211","OPU","1","","",DT_ME21N_1211_TABLECELL_OPU_0,False)    
Call SetTableData("SAPLMEGUITC_1211","Matl Group","1","","",DT_ME21N_1211_TABLECELL_MDSE_CAT_0,False)     
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False)     

Call PressEnter()   
wait 5
Call TakeScreenShot()
Call SetFocusGuiLabel("20010501",11,72,False) 
wait 5
Call PressEnter()

Call SetComboByKey("MEACCT1200-KNTTP",DT_ME21N_1200_ACCASSCAT)     
Call SetTextbox("G/L Account","MEACCT1100-SAKTO","",DT_GL_ACCOUNT,False)     
Call SetTextbox("Business Area","COBL-GSBER","",DT_ME21N_1101_BUSINESS_AREA,False)     
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_ME21N_1101_COST_CENTER,False)     

'''Call ClickButton("btn\[11\]",False)    

Call ClickButton("Save   \(Ctrl\+S\)",False)    
'''Call ClickButton("SPOP-VAROPTION1",True)     
Call ClickButton("Save",True)     
Call GetStatusBar("item2","DT_DOC_NUM_OUTPUT")

Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NUM_OUTPUT",DT_DOC_NUM)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
