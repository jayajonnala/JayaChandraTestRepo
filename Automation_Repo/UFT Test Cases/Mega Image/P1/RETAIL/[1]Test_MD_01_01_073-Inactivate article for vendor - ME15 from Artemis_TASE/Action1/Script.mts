
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_073-Inactivate article for vendor - ME15 from Artemis
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

gstrTestCaseName = "Test_MD_01_01_073-ME15 from Artemis"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_MD_01_01_073-Inactivate article for vendor - ME15 - inactivate info record from artemis - complete info only.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''--------TransactionCode-MM43----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM43_0100_ARTICLE,false)
Call SetTextbox("Purchasing Org\.","RMMW1-EKORG","",DT_MM43_0100_PURCHASING_ORG,false)
'''Call SetTextbox("Vendor","RMMW1-LIFNR","",DT_MM43_0100_VENDOR,false)
Call SetTextboxNoLabel("RMMW1-LIFNR","",DT_MM43_0100_VENDOR,false)

Call ClickButton("Deselect All   \(Shift\+F7\)",False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description", DT_SELECT_VIEWS_PURCHASING, false)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("No",True)
Call ClickButton("IMPORTDATEN_PUSH",False)   ''For InfoRecord
Call GetTextboxValue("EINA-INFNR", "", "DT_MM43_0101_CHECK_TEXT_OF_INFO_RECORD_OUTPUT", False)

'''--------TransactionCode-ME15----------''''
Call SetTcode(DT_MM43_0101_OKCD)     
Call PressEnter()     
Call TakeScreenShot

'''Call SetTextbox("Vendor","EINA-LIFNR","",DT_MM43_0100_VENDOR_OCC1,false)
Call SetTextboxNoLabel("EINA-LIFNR","",DT_MM43_0100_VENDOR_OCC1,false)
Call SetTextbox("Article","EINA-MATNR","",DT_MM43_0100_ARTICLE_OCC1,false)
Call SetTextbox("Purchasing Org.","EINE-EKORG","",DT_MM43_0100_PURCHASING_ORG_OCC1,false)

Call SelectRadioButton("RM06I-NORMB", "Standard", False)
Call PressEnter()     
Call TakeScreenShot

Call SelectCheckbox("EINA-LOEKZ", 0, DT_MM43_0104_COMPLETE_INFO_RECORD, False)
Call SelectCheckbox("EINE-LOEKZ", 1, DT_MM43_0104_PURCH_ORG_DATA, False)

Call ClickButton("Save   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_MM43_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Purchasing info record "&DT_MM43_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" RP01   changed")

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
