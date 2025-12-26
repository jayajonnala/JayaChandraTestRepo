
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_013-Create Display Article - SAP_P2     
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

gstrTestCaseName = "Test_MD_01_01_013-Create Display Article - SAP_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_MD_01_01_013-Create Display Article - SAP_P2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''--------TransactionCode-ZMDAS_WSL11----------''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Assortment","S_FILIA-LOW","",DT_ZMDAS_WSL11_1000_ASSORTMENT,False)
Call SetTextbox("Article","S_ARTNR-LOW","",DT_ZMDAS_WSL11_1000_ARTICLE,False)
Call SetTextbox("Valid From","P_DATAB","",ConvertDate(DT_ZMDAS_WSL11_1000_VALID_FROM),False)
Call SetTextbox("Valid To","P_DATBI","",ConvertDate(DT_ZMDAS_WSL11_1000_VALID_TO),False)
Call SetTextbox("Customer No\. - Site","S_LOCNR-LOW","",DT_SITE,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyGridCellContentbyName("shell", 1, "ARTNR", "", DT_ZMDAS_WSL11_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ARTNR)
Call VerifyGridCellContentbyName("shell", 1, "MAKTX", "", lcase(DT_ZMDAS_WSL11_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MAKTX))
Call VerifyGridCellContentbyName("shell", 1, "FILIA", "", DT_ZMDAS_WSL11_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_FILIA)
'
''''''--------TransactionCode-MM42----------''''
'''
'''Call SetTcode(DT_ZMDAS_WSL11_0500_OKCD)     
'''Call PressEnter()     
'''Call SetTextbox("Article","RMMW1-MATNR","",DT_ZMDAS_WSL11_0100_ARTICLE,False)
'''Call SetTextbox("Sales Org.","RMMW1-VKORG","",DT_ZMDAS_WSL11_0100_SALES_ORG,False)
'''Call SetTextbox("Distr. Channel","RMMW1-VTWEG","",DT_ZMDAS_WSL11_0100_DISTR_CHANNEL,False)
'''Call ClickButton("Deselect All   \(Shift\+F7\)",False)
'''Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description", "Sales", False)
'''Call TakeScreenShot
'''Call PressEnter()
'''Call TakeScreenShot
'''Call SetTextbox("Sales unit","MVKE-VRKME","","",False)
'''Call TakeScreenShot
'''Call ClickButton("MVKE_PUSH",False)  ''other sales data
'''Call TakeScreenShot
'''Call SetTextbox("Article pricing grp","MVKE-KONDM","",DT_ZMDAS_WSL11_2152_ARTICLE_PRICING_GRP,False)
'''Call TakeScreenShot
'''Call ClickButton("btn\[39\]",False)  ''main data
'''Call TakeScreenShot
'''Call ClickButton("btn\[13\]",False)
'''Call SetTextbox("Distr. Channel","RMMW1-VTWEG","",DT_ZMDAS_WSL11_0081_DISTR_CHANNEL,True)
'''Call SetTextbox("Sales Org.","RMMW1-VKORG","",DT_ZMDAS_WSL11_0081_SALES_ORG,True)
'''Call SetTextbox("Sales unit ","RMMW1-VRKME","",DT_ZMDAS_WSL11_0081_SALES_UNIT,True)
'''Call TakeScreenShot
'''Call ClickButton("Continue   \(Enter\)",True)
'''Call TakeScreenShot
'''Call SetTextbox("Final price","CALP-ENDPR","",DT_ZMDAS_WSL11_2233_FINAL_PRICE,False)
'''   
'''Call TakeScreenShot
'''Call PressEnter()   
'''Call TakeScreenShot
'''Call ClickButton("Save   \(Ctrl\+S\)",False)
'''Call TakeScreenShot
'''Call VerifyStatusBar(DT_ZMDAS_WSL11_0100_CHECK_TEXT_OF_STATUSBAR)
'''
   
'''----------------------Tcode VKP5 - Setting sales price steps is changed from MM42 to VKP5 ----------------------------
Call SetTcode(DT_ZMDAS_WSL11_0500_OKCD_1)     
Call PressEnter()
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("btn\[17\]",False)
Call SetTextbox("Variant","V-LOW","",DT_VKP5_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",True)
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Article","S_MATNR-LOW","",DT_ZMDAS_WSL11_0100_ARTICLE,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)

Call TakeScreenShot
Call SetGridData("", 1, "ENDPR", DT_ZMDAS_WSL11_2233_FINAL_PRICE, False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call GetStatusBar("MessageType", "DT_VKP5_1000_GET_MESSAGE_TYPE_OUTPUT")
Call GetStatusBar("item1", "DT_ZMDAM_BOM_REPORT_1000_GET_PRICING_DOCUMENT_OUTPUT")
Call VerifyStatusBar("Data saved; pricing document "&DT_ZMDAM_BOM_REPORT_1000_GET_PRICING_DOCUMENT_OUTPUT&" created")
Call TakeScreenShot
'''--------TransactionCode-VKP2----------''''

Call SetTcode(DT_ZMDAS_WSL11_0100_OKCD)     
Call PressEnter()     
Call SetTextbox("Article","S_MATNR-LOW","",DT_ZMDAS_WSL11_1000_ARTICLE_OCC1,False)
Call SetTextbox("Sales organization","S_VKORG-LOW","",DT_ZMDAS_WSL11_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution channel","S_VTWEG-LOW","",DT_ZMDAS_WSL11_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Validity Period","S_DATUM-LOW","",ConvertDate(DT_ZMDAS_WSL11_1000_VALIDITY_PERIOD),False)
Call SetTextbox("to","S_DATUM-HIGH","",ConvertDate(DT_ZMDAS_WSL11_1000_TO),False)
Call TakeScreenShot
Call PressEnter()   
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyGridCellContent("Sales Price Conditions", 1, "MATNR", 0, DT_ZMDAS_WSL11_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call VerifyGridCellContent("Sales Price Conditions", 1, "KBETR", 0, DT_ZMDAS_WSL11_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KBETR)
Call VerifyGridCellContent("Sales Price Conditions", 1, "KSCHL", 0, DT_ZMDAS_WSL11_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)

Call LogOff()

Call FinalStatus ()








'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [8,3437477]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


