

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04-02-Article MD- Fresh article BE creation via Matcopy
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

gstrTestCaseName = "Test_04-02-Article MD- Fresh article BE creation via Matcopy"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'
''//-----------------------------------ZMDAM_WRFMATCOPY -----------------------------------
'
Call ClickButton("Load selection screen variant   \(Shift\+F5\)",False)
Call SetTextbox("Screen Variant","ZMDAM_MATCOPY_SV-MATCOPY_VARIANT","",DT_ZMDAM_WRFMATCOPY_0201_SCREEN_VARIANT,True) 
Call TakeScreenShot()
Call ClickButton("OK   \(F2\)",True)
Call SetTextbox("Article","ZMDAM_S_MATCOPY_SSCR-P_RMATN","",DT_ZMDAM_WRFMATCOPY_0050_ARTICLE,False) 
Call PressEnter()     ' 
Call TakeScreenShot()
Call ClickButton("Key Filters",False)
Call SelectRadioButton("SPOPLI-SELFLAG","Hierarchy ID",True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Deselect all   \(F6\)",True)
Call SelectRowGuiGridbyRowNo("Reduction Level",0,3,True)
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

''no need of using the GTIn for the TC so commenting the line of code.

''Call SetGridData("Variant Filters",1,"EAN11",DT_ZMDAM_WRFMATCOPY_0700_GRIDCELL_0_GTIN,True)
''Call SetGridData("Variant Filters",2,"EAN11",DT_ZMDAM_WRFMATCOPY_0700_GRIDCELL_1_GTIN,True)
''Call TakeScreenShot()
''Call SetGridData("Variant Filters",1,"MAIN_EAN","ON",True)
''Call SetGridData("Variant Filters",2,"MAIN_EAN","ON",True)
''Call SelectRowGuiGridbyRowNo("Variant Filters",0,2,True)

Call ClickButton("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
wait 60
Call TakeScreenShot()
Call GetGridContent("",0,"MESSAGE_V1",3,"<NA>","<NA>","DT_ZMDAM_WRFMATCOPY_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MESSAGE_V1_OUTPUT")
Call VerifyGridCellContent("",3,"Message text",0,"Article "&DT_ZMDAM_WRFMATCOPY_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MESSAGE_V1_OUTPUT&" created")

Call LogOff()
Call FinalStatus ()



'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [12,8640008]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


