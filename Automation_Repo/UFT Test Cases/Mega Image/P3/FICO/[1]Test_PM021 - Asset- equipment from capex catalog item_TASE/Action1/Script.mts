
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PM021 - Asset- equipment from capex catalog item
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


gstrTestCaseName = "Test_PM021 - Asset- equipment from capex catalog item"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

''gstrInputExcelFilePathAndName="C:\Users\rsara\Desktop\TASEWork\Data\TASE_DT_PM021 - Asset- equipment from capex catalog item SC.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
''DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''--------TransactionCode-AS03----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS03_0100_ASSET,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS03_0100_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot

Call SelectTab("TABSTRIP100", "Allocations", False)

Call GetTableCellData("SAPLAISTTC_EQUI", "Equipment number", 1, "<NA>", "<NA>", "EQUIPMENT_NUMBER_OUTPUT", False)

''---------/nIE03---------------------
Call SetTcode(DT_AS03_1000_OKCD)     
Call PressEnter()     
Call TakeScreenShot


'''''Call SetTextbox("Equipment","RM63E-EQUNR","",EQUIPMENT_NUMBER_OUTPUT,False)
Call SetTextbox("Equipment","RM63E-EQUNR","",DT_AS03_0100_EQUIPMENT,False)

Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot

Call VerifyTextBoxContent("Status", "ITOBATTR-STTXT", 0, DT_AS03_1526_CHECK_TEXT_OF_STATUS, False)

Call SelectTab("TABSTRIP", "Location", False)
Call TakeScreenShot
Call VerifyTextBoxContent("MaintSite", "ITOB-SWERK", 0, DT_AS03_1050_CHECK_TEXT_OF_MAINTSITE, False)
Call VerifyTextBoxContent("Sort field", "ITOB-EQFNR", 0, "", False)

Call SelectTab("TABSTRIP", "Organization", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Company Code", "ITOB-BUKRS", 0, DT_AS03_1052_CHECK_TEXT_OF_COMPANY_CODE, False)
Call VerifyTextBoxContent("Business Area", "ITOB-GSBER", 0, DT_AS03_1052_CHECK_TEXT_OF_BUSINESS_AREA, False)
Call VerifyTextBoxContent("Asset", "ITOB-ANLNR", 0, DT_AS03_1052_CHECK_TEXT_OF_ASSET, False)
Call VerifyTextBoxContent("Cost Center", "ITOB-KOSTL", 0, DT_AS03_1052_CHECK_TEXT_OF_COST_CENTER, False)
Call VerifyTextBoxContent("Planning site", "ITOB-IWERK", 0, DT_AS03_1062_CHECK_TEXT_OF_PLANNING_SITE, False)
Call VerifyTextBoxContent("Main WorkCtr", "ITOBATTR-GEWRK", 0, DT_AS03_1062_CHECK_TEXT_OF_MAIN_WORKCTR, False)

Call SelectTab("TABSTRIP", "Structure", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Functional loc\.", "ITOB-TPLNR", 0, DT_AS03_1065_CHECK_TEXT_OF_FUNCTIONAL_LOC, False)

Call SelectTab("TABSTRIP", "Warranties", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Warranty Start", "WCHECK_V_H-GWLDT_O", 0, "", False)
Call VerifyTextBoxContent("Warranty end", "WCHECK_V_H-GWLEN_O", 0, "", False)					

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

