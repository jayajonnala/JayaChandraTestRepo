

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_03.03.01.02.01 Manage Function Location - Creation Functional Ar
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

gstrTestCaseName = "Test_03.03.01.02.01 Manage Function Location - Creation Functional Ar"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''----------------------Tcode IL01----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(DT_INCREMENT+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


'Enter Details
Call SetTextbox("FunctLocCat\.","IFLO-FLTYP","",DT_IL01_1110_FUNCTLOCCAT,False) 
Call SetTextbox("StrIndicator","RILO0-TPLKZ","",DT_IL01_1110_STRINDICATOR,False) 
Call SetTextbox("Functional Loc\.","IFLOS-STRNO","",DT_IL01_1110_FUNCTIONAL_LOC,False)
Call TakeScreenShot()
Call PressEnter()  


'Call SetTextbox("Description","IFLO-PLTXT","",DT_IL01_2100_DESCRIPTION,False) 
Call SetSpecialTextbox("Description","IFLO-PLTXT",0,DT_IL01_2100_DESCRIPTION,False)

'Navigate to Location Tab
Call SelectTab("TABSTRIP","Location",False)
Wait(1)

Call SetTextbox("ABC indic\.","ITOB-ABCKZ","",DT_IL01_1050_ABC_INDIC,False) 
Call TakeScreenShot()
Call PressEnter()  

'Navigate to Organization Tab
Call SelectTab("TABSTRIP","Organization",False)
Wait(1)
Call TakeScreenShot()

'Enter Details
Call SetTextbox("Planner group","ITOB-INGRP","",DT_IL01_1062_PLANNER_GROUP,False)
Call SetTextbox("Main WorkCtr","ITOBATTR-GEWRK","",DT_IL01_1062_MAIN_WORKCTR,False)
Call SetTextboxNoLabel("ITOBATTR-WERGW",0,DT_IL01_1062_ITOBATTRWERGW,False)
Call TakeScreenShot()
Call PressEnter()


Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S")

Call GetStatusBar("item1","DT_FUNCLOC_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_SUCCESSMESSAGE)

''----------------------Tcode IL03----------------------------

Call SetTcode(DT_IL01_1110_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_IL01_1110_OKCD)

Call PressEnter()
Call PressEnter()

Call VerifyTextBoxContent("Functional loc\.","IFLO-TPLNR",0,DT_IL01_2100_CHECK_TEXT_OF_FUNCTIONAL_LOC,False)
'Call VerifyTextBoxContent("Description","IFLO-PLTXT",0,DT_IL01_2100_CHECK_TEXT_OF_DESCRIPTION,False)
Call VerifyTextBoxContent("Status","RILO0-STTXT",0,DT_IL01_2100_CHECK_TEXT_OF_STATUS,False)


'Navigate to Location Tab
Call SelectTab("TABSTRIP","Location",False)
Wait(1)

Call VerifyTextBoxContent("MaintSite","ITOB-SWERK",0,DT_IL01_1050_CHECK_TEXT_OF_MAINTSITE,False)
Call VerifyTextBoxContent("ABC indic\.","ITOB-ABCKZ",0,DT_IL01_1050_CHECK_TEXT_OF_ABC_INDIC,False)


'Navigate to Organization Tab
Call SelectTab("TABSTRIP","Organization",False)
Wait(1)
Call TakeScreenShot()


Call VerifyTextBoxContent("Company Code","ITOB-BUKRS",0,DT_IL01_1052_CHECK_TEXT_OF_COMPANY_CODE,False)
Call VerifyTextBoxContent("Business Area","ITOB-GSBER",0,DT_IL01_1052_CHECK_TEXT_OF_BUSINESS_AREA,False)
Call VerifyTextBoxContent("Planning site","ITOB-IWERK",0,DT_IL01_1062_CHECK_TEXT_OF_PLANNING_SITE,False)
Call VerifyTextBoxContent("Planner group","ITOB-INGRP",0,DT_IL01_1062_CHECK_TEXT_OF_PLANNER_GROUP,False)
Call VerifyTextBoxContent("Main WorkCtr","ITOBATTR-GEWRK",0,DT_IL01_1062_CHECK_TEXT_OF_MAIN_WORKCTR,False)
'Call VerifyTextBoxContent("","ITOBATTR-WERGW",0,DT_IL01_1062_CHECK_TEXT_OF_ITOBATTRWERGW,False)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
