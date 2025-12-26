

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06DCIM03_013_Pallet_File_Quantity_Deficit_in_status_10_with_tie_P03
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


gstrTestCaseName = "Test_06DCIM03_013_10_with_tie_P03"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Article","MATNR-LOW","",DT_MB51_1000_ARTICLE,False)
Call SetTextbox("Site","WERKS-LOW","",DT_MB51_1000_SITE,False)
Call SetTextbox("Movement type","BWART-LOW","",DT_MB51_1000_MOVEMENT_TYPE,False)
Call SetTextbox("Document Header Text","BKTXT-LOW","",DT_MB51_1000_DOCUMENT_HEADER_TEXT,False)
Call SetTextbox("Document Date","BLDAT-LOW","",ConvertDate(DT_MB51_1000_DOCUMENT_DATE),False)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call ClickButton("Detail List   \(Ctrl\+Shift\+F12\)",False)
Call TakeScreenShot()
Call SelectMenuBar("Settings;Layout;Choose...")

Call ClickCellGuiGrid("",0,"Layout description","","Layout description",DT_MB51_0501_GRIDCELL_5_LAYOUT_DESCRIPTION,True)
Call VerifyStatusBarMessageType("S")
Call VerifyGridCellContent("",1,"MATNR",0,DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call VerifyGridCellContent("",1,"ERFMG",0,DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ERFMG)
Call VerifyGridCellContent("",1,"LGORT",0,DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LGORT)

''-------------------------------MMBE-----------------------------


Call SetTcode(DT_MB51_0500_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MB51_0500_OKCD)

Call SetTextbox("Article","MS_MATNR-LOW","",DT_MB51_1000_ARTICLE_OCC1,False)
Call SetTextbox("Site","MS_WERKS-LOW","",DT_MB51_1000_SITE_OCC1,False)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call VerifyNodeTextGuiTree(0,DT_MB51_0300_CHECK_TEXT_OF_TREE_FULL_NODE)

Call LogOff()
Call FinalStatus ()


