
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_090-Change EAN from article A to article B_p1
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

gstrTestCaseName = "Test_MD_01_01_090- A to article B_p1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

''gstrInputExcelFilePathAndName="C:\Users\rsara\Desktop\TASEWork\Data\TASE_DT_MD_01_01_090-Change EAN from article A to article B_p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

''DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()




'''--------TransactionCode-MM42----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE,False)
Call SetTextbox("Sales Org\.","RMMW1-VKORG","",DT_MM42_0100_SALES_ORG,False)
Call SetTextbox("Distr\. Channel","RMMW1-VTWEG","",DT_MM42_0100_DISTR_CHANNEL,False)
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100", 1, False)
'''''''' SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description", "Basic Data", False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot

Call GetTableCellData("SAPLMGD2TC_ME_8022", "GTIN", 1, "", "","DT_MM42_8022_TABLECELL_GTIN_0_OCC1_OUTPUT", False)

Call SetTableData("SAPLMGD2TC_ME_8022", "GTIN", 1, "", "", DT_MM42_8022_TABLECELL_GTIN_0, False)

Call SetTableData("SAPLMGD2TC_ME_8022", "Ct", 1, "", "", DT_MM42_8022_TABLECELL_CT_0, False)

Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButtonIfExist("Yes",True)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButtonIfExist("Yes",True)
Call VerifyStatusBar("Article "&DT_MM42_0100_ARTICLE&" changed")
Call TakeScreenShot

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE_OCC1,False)
Call SetTextbox("Sales Org\.","RMMW1-VKORG","",DT_MM42_0100_SALES_ORG_OCC1,False)
Call SetTextbox("Distr\. Channel","RMMW1-VTWEG","",DT_MM42_0100_DISTR_CHANNEL_OCC1,False)
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100", 1, False)
Call TakeScreenShot
Call PressEnter() 

Call SetTableData("SAPLMGD2TC_ME_8022", "GTIN", 1, "", "", DT_MM42_8022_TABLECELL_GTIN_0_OCC1_OUTPUT, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Ct", 1, "", "", DT_MM42_8022_TABLECELL_CT_0_OCC1, False)
	
Call PressEnter() 
Call TakeScreenShot
Call ClickButtonIfExist("Yes",True)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButtonIfExist("Yes",True)
Call PressEnter() 
Call TakeScreenShot
Call ClickButtonIfExist("Yes",True)
Call VerifyStatusBar("Article "&DT_MM42_0100_ARTICLE_OCC1&" changed")
Call TakeScreenShot
''''''''Call VerifyTableCellContent(1, "GTIN", "SAPLMGD2TC_ME_8022", DT_MM41_8022_TABLECELL_GTIN_0)

Call LogOff()

Call FinalStatus ()









'''//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''
'''Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'''Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'''Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 
''
'''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''
''
'''// ---- Script Generated in [0] Minutes , [13,4062483]  Seconds ---- //
'''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''' ................NOTE: 
'''.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'''.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'''.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'''.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
''' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''

