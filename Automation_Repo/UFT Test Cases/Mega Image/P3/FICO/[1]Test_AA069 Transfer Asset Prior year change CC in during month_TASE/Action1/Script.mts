
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA069 Transfer Asset Prior year change CC in during month
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


gstrTestCaseName = "Test_AA069 Transfer Asset Prior year change CC in during month"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

''gstrInputExcelFilePathAndName="C:\Users\rsara\Desktop\TASEWork\Data\TASE_DT_AA069 Transfer Asset Prior year change CC in during month.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
''DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()




''--------TransactionCode-AS02----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


 Call SetTextbox("Asset","ANLA-ANLN1", "", DT_AS02_0100_ASSET, False)
 Call SetTextbox("Sub-number","ANLA-ANLN2", "", DT_AS02_0100_SUBNUMBER, False)
 Call SetTextbox("Company Code","ANLA-BUKRS", "", DT_AS02_0100_COMPANY_CODE, False)
 Call TakeScreenShot
 Call PressEnter() 
Call TakeScreenShot

Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
''SET Default values  for cost center, profit center fields  - RS000001
Call SetTextbox("Cost Center","ANLZ-KOSTL", "", DT_COST_CENTER_DEFAULT, False)
Call SetTextbox("Profit Center","ANLZ-PRCTR", "", DT_PROFIT_CENTER_DEFAULT, False)

Call ClickButtonIfExist("More Intervals", False)
Call TakeScreenShot
Call PressEnter() 
Call ClickButtonIfExist("Insert Row", False)
Call TakeScreenShot
Call SetTextbox("From-date of new interval","ANLZ-ADATU", "", DT_AS02_3010_FROMDATE_OF_NEW_INTERVAL, True)
Call ClickButtonIfExist("Yes   \(Enter\)", True)
Call TakeScreenShot

Call SetTableDataNoRef("SAPLAISTTIME", "Cost Center", 1, DT_AS02_3000_TABLECELL_COST_CENTER_0, False)
Call SetTableDataNoRef("SAPLAISTTIME", "Profit Center", 1, DT_AS02_3000_TABLECELL_PROFIT_CENTER_0, False)



Call PressEnter() 
Call ClickButtonIfExist("Save   \(Ctrl\+S\)", False)
Call PressEnter() 
Call PressEnter() 

Call GetStatusBar("item1", "DT_AS02_ASSET_OUTPUT")

Call VerifyStatusBar("The asset "&DT_AS02_ASSET_OUTPUT&" 0 is changed")
Call TakeScreenShot

''--------TransactionCode-/NEX ----------''''
''Call SetTcode(DT_AS02_0100_OKCD)
''Call PressEnter()     
''Call TakeScreenShot
''Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)



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

