
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "Test_Maintain Stock_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
SAPGuiUtil.OpenConnection("R1E - SAP RETAIL Pre-Production EUROPE")
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'-------------------------------MI10---------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call TakeScreenShot
Call SetTextbox("Site","IKPF-WERKS","",DT_MI10_0700_SITE,False)    
Call SetTextbox("Storage Location","IKPF-LGORT","",DT_MI10_0700_STORAGE_LOCATION,false)
'Call SetTextbox("Storage Location","IKPF-LGORT","","0001",false) 
Call PressEnter()
Call TakeScreenShot()
'DT_MI10_0731_ARTICLE_DESCRIPTION changed to DTARTICLE
Call SetTextbox("Material Description","ISEG-MATNR",0,DTARTICLE,false)  
'DT_MI10_0731_ISEGERFMG
call SetTextboxNoLabel("ISEG-ERFMG",0,DT_MI10_0731_ISEGERFMG,false)  'quantity coumn
call SetTextboxNoLabel("ISEG-BSTAR",0,DT_MI10_0731_ISEGBSTAR,false)  'sty column
Call TakeScreenShot()
Call SetTextbox("Material Description","ISEG-MATNR",1,DTARTICLE_OCC,false)  
call SetTextboxNoLabel("ISEG-ERFMG",1,DT_MI10_0731_ISEGERFMG_OCC1,false)  'quantity coumn
call SetTextboxNoLabel("ISEG-BSTAR",1,DT_MI10_0731_ISEGBSTAR_OCC1,false)  'sty column
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("btn\[11\]",false)
Call TakeScreenShot
'Call PressEnter()
'Call TakeScreenShot

call GetStatusBar("item1","DT_Invoicedoc_Output")
'Call VerifyStatusBar(DT_Invoicedoc_Output)

Call TakeScreenShot
Call LogOff()
Call FinalStatus ()

