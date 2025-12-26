
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0132-Production of juice in RW29_P1_Get MAP and planning     
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

gstrTestCaseName = "Test_P2P_01_01_0132_P1_Get MAP and planning"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_0132-Production of juice in RW29_P1_Get MAP and planning.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


'''--------TransactionCode-MM43----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM43_0100_ARTICLE,False)
Call SetTextbox("Purchasing Org.","RMMW1-EKORG","",DT_MM43_0100_PURCHASING_ORG,False)
'for selecting basic view we added below line
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100", 1, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call VerifyTextBoxContent("Material Type", "MARA-MTART", "", DT_MM43_2001_CHECK_TEXT_OF_ARTICLE_TYPE, False)

Call SelectTab("TABSPR1", "Logistics: DC", False)

Call TakeScreenShot
Call ClickButton("Switch area of validity   \(Shift\+F1\)",False)
Call SetTextbox("Distr. center","RMMW1-VZWRK","",DT_MM43_0081_DISTR_CENTER,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Accounting",False)
Call TakeScreenShot
Call GetTextboxValue("MBEW-VERPR", "Moving price", "DT_MM43_2802_CHECK_TEXT_OF_MOVING_PRICE_OUTPUT", False)


'''--------TransactionCode-MM43----------''''

Call SetTcode(DT_MM43_3000_OKCD)     
Call PressEnter()  

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM43_0100_ARTICLE_OCC1,False)
'for selecting basic view below line is added
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100", 1, False)
Call TakeScreenShot
Call PressEnter()  
Call SelectTab("TABSPR1", "Logistics: DC", False)
Call TakeScreenShot
Call ClickButton("Switch area of validity   \(Shift\+F1\)",False)
Call SetTextbox("Distr. center","RMMW1-VZWRK","",DT_MM43_0081_DISTR_CENTER_OCC1,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Accounting",False)
Call TakeScreenShot
Call GetTextboxValue("MBEW-VERPR", "Moving price", "DT_MM43_2802_CHECK_TEXT_OF_MOVING_PRICE_OCC1_OUTPUT", False)

'''--------TransactionCode-MM43----------''''

Call SetTcode(DT_MM43_3000_OKCD_OCC1)     
Call PressEnter()  

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM43_0100_ARTICLE_OCC2,False)
Call SetTextbox("Purchasing Org.","RMMW1-EKORG","",DT_MM43_0100_PURCHASING_ORG_OCC1,False)
'for selecting basic view below line is added
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100", 1, False)
Call TakeScreenShot
Call PressEnter()  
Call SelectTab("TABSPR1", "Logistics: DC", False)
Call TakeScreenShot
Call ClickButton("Switch area of validity   \(Shift\+F1\)",False)
Call SetTextbox("Distr. center","RMMW1-VZWRK","",DT_MM43_0081_DISTR_CENTER_OCC2,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Accounting",False)
Call TakeScreenShot
Call GetTextboxValue("MBEW-VERPR", "Moving price", "DT_MM43_2802_CHECK_TEXT_OF_MOVING_PRICE_OCC2_OUTPUT", False)

'''--------TransactionCode-MM43----------''''

Call SetTcode(DT_MM43_3000_OKCD_OCC2)     
Call PressEnter()  

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM43_0100_ARTICLE_OCC3,False)
Call SetTextbox("Purchasing Org.","RMMW1-EKORG","",DT_MM43_0100_PURCHASING_ORG_OCC2,False)
'for selecting basic view below line is added
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100", 1, False)
Call TakeScreenShot
Call PressEnter()  
Call SelectTab("TABSPR1", "Logistics: DC", False)
Call TakeScreenShot
Call ClickButton("Switch area of validity   \(Shift\+F1\)",False)
Call SetTextbox("Distr. center","RMMW1-VZWRK","",DT_MM43_0081_DISTR_CENTER_OCC3,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Accounting",False)
Call TakeScreenShot
Call GetTextboxValue("MBEW-VERPR", "Moving price", "DT_MM43_2802_CHECK_TEXT_OF_MOVING_PRICE_OCC3_OUTPUT", False)

'''--------TransactionCode-MM43----------''''

Call SetTcode(DT_MM43_3000_OKCD_OCC3)     
Call PressEnter()  

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM43_0100_ARTICLE_OCC4,False)
Call SetTextbox("Purchasing Org.","RMMW1-EKORG","",DT_MM43_0100_PURCHASING_ORG_OCC3,False)
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100", 1, False)
Call TakeScreenShot
Call PressEnter()  
Call SelectTab("TABSPR1", "Logistics: DC", False)
Call TakeScreenShot
Call ClickButton("Switch area of validity   \(Shift\+F1\)",False)
Call SetTextbox("Distr. center","RMMW1-VZWRK","",DT_MM43_0081_DISTR_CENTER_OCC4,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Accounting",False)
Call TakeScreenShot
Call GetTextboxValue("MBEW-VERPR", "Moving price", "DT_MM43_2802_CHECK_TEXT_OF_MOVING_PRICE_OCC4_OUTPUT", False)
'
'''--------TransactionCode-MM43----------''''

Call SetTcode(DT_MM43_0110_OKCD)     
Call PressEnter() 

Call SetTextbox("Planned Order Profile","RM61P-PASCH","",DT_PLANNED_ORDER_PROFILE,False)
Call SetTextbox("Planned Order","RM61P-PLREF","","",False)
Call PressEnter()
Call SetTextbox("Material","PLAF-MATNR","",DT_MM43_0110_ARTICLE,False)
Call SetTextbox("Planning Plant","PLAF-PLWRK","",DT_MM43_0110_PLANNING_SITE,False)
Call SetTextbox("Production Version","PLAF-VERID","",DT_MM43_0810_PRODUCTION_VERSION,False)
Call SetTextbox("Storage Location","PLAF-LGORT","",DT_MM43_0810_STORAGE_LOCATION,False)
Call SetTextbox("Production Plant","PLAF-PWWRK","",DT_MM43_0810_PRODUCTION_SITE,False)
Call TakeScreenShot
Call SetTextbox("End","PLAF-PEDTR","",ConvertDate(DT_MM43_0802_END),False)
Call SetTextbox("Order quantity","PLAF-GSMNG","",DT_MM43_0802_ORDER_QUANTITY,False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot

For Iterator = 1 To 4 Step 1
 Stausbartext = SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiStatusBar(guiStatusBar).GetROProperty("Text")
	If Instr(1,Stausbartext,"will be created") > 0 then 
		Exit For		
		Else
		Call PressEnter()
	End If
Next

Call GetTextStatusBar("DT_GET_ORDER_MESSAGE_OUTPUT")

Call VerifyStatusBar(LCase(DT_GET_ORDER_MESSAGE_OUTPUT))


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



