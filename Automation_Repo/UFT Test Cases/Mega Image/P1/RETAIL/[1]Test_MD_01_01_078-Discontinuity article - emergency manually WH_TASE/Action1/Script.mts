
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_078-Discontinuity article - emergency manually WH
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

gstrTestCaseName = "Test_MD_01_01_078-Discontinuity article - emergency manually WH"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\ssahoo\Desktop\TASEWork\Data\P1-MI\TASE_DT_MD_01_01_078-Discontinuity article - emergency manually WH.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''--------TransactionCode-ZMDAS_WSL11----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Assortment","S_FILIA-LOW","",DT_ZMDAS_WSL11_1000_ASSORTMENT,False)
Call SetTextbox("Article","S_ARTNR-LOW","",DT_ZMDAS_WSL11_1000_ARTICLE,False)
Call SetTextbox("Valid From","P_DATAB","",ConvertDate(DT_ZMDAS_WSL11_1000_VALID_FROM),False)
Call SetTextbox("Customer No\. - Site","S_LOCNR-LOW","",DT_ZMDAS_WSL11_1105_CUSTOMER_NO__SITE,False)

Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyGridCellContentbyName("shell", 1, "LOCNR", "", DT_ZMDAS_WSL11_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOCNR)

''--------TransactionCode-WSE6----------''''

Call SetTcode(DT_ZMDAS_WSL11_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot
'Call CheckTCodeScreen(DT_ZMDAS_WSL11_0500_OKCD)

Call SetTextbox("Assortment","S_ASORT-LOW","",DT_ZMDAS_WSL11_1110_ASSORTMENT,false)
Call SetTextbox("Article","S_MATNR-LOW","",DT_ZMDAS_WSL11_1110_ARTICLE,false)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call SelectRowGuiGrid("Article discontinuation for assortmnt", 0, "Assortment", DT_ZMDAS_WSL11_1110_ASSORTMENT, False)


Call Click204ButtonToolBar("DELE","")
''Call VerifyGridCellContentbyName("shell",1,"Assortment status","",DT_ZMDAS_WSL11_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SSTAT)


'''--------TransactionCode-/nZMDAS_WSL11----------''''

Call SetTcode(DT_ZMDAS_WSL11_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Assortment","S_FILIA-LOW","",DT_ZMDAS_WSL11_1000_ASSORTMENT_OCC1,False)
Call SetTextbox("Article","S_ARTNR-LOW","",DT_ZMDAS_WSL11_1000_ARTICLE_OCC1,False)
Call SetTextbox("Valid From","P_DATAB","",ConvertDate(DT_ZMDAS_WSL11_1000_VALID_FROM_OCC1),False)
Call SetTextbox("Customer No\. - Site","S_LOCNR-LOW","",DT_ZMDAS_WSL11_1105_CUSTOMER_NO__SITE_OCC1,False)

Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectColumnGuiGrid("", 0, "Customer No. - Site", False)

Call ClickButton("Set filter   \(Ctrl\+F5\)",False)

Call SetTextbox("Customer No. - Site","%%DYN001-LOW","",DT_ZMDAS_WSL11_1110_ASSORTMENT,True)
Call PressEnter()  
'Call verifyNoRowExistsGrid("shell",0)

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




