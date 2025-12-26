'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_PRE_P2P_01_01_072-Physical inventory Non VBS -SAP only RW04_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 9th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE_P2P_01_01_072 only RW04_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_PRE_P2P_01_01_072-Physical inventory  Non VBS -SAP only  RW04_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenari
''''----------------------Login----------------------------

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''----------------------Tcode MMBE----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Article","MS_MATNR-LOW","",DT_MMBE_1000_ARTICLE,False)
Call SetTextbox("Site","MS_WERKS-LOW","",DT_MMBE_1000_SITE,False)
Call SetTextbox("Storage location","MS_LGORT-LOW","",DT_MMBE_1000_STORAGE_LOCATION,False)
Call SetTextbox("Batch","MS_CHARG-LOW","","",False)
Call SetTextbox("Special Stock Indicator","SOBKZ-LOW","","",False)
Call SetTextbox("Display version","VERNU","",DT_MMBE_1000_DISPLAY_VERSION,False)
Call SetTextbox("Display Unit of Measure","MEINH","","",False)
'Capture the screenshot
Call TakeScreenShot()
Call SelectCheckbox("KZNUL","1","OFF",False)
Wait(2)
Call ClickButton("Execute   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

If DT_MMBE_1000_STORAGE_LOCATION = "0001" Then
	Call ActivateItemGuiTree("","Full;RO02 Mega Image S.R.L.;RW22 WH STEFANESTI F&V;0001 Healthy stock","0001 Healthy stock")
End If

If DT_MMBE_1000_STORAGE_LOCATION = "0007" Then
	Call ActivateItemGuiTree("","Full;RO02 Mega Image S.R.L.;RW22 WH STEFANESTI F&V;0007 Blocked stock","0007 Blocked stock")
End If

If DT_MMBE_1000_STORAGE_LOCATION = "0008" Then
	Call ActivateItemGuiTree("","Full;RO02 Mega Image S.R.L.;RW22 WH STEFANESTI F&V;0008 Return to Vendor","0008 Return to Vendor")
End If

Call GetCellDataGuiGridPopupByRefTwoColumns("SLoc","","Stock","Stock Type",DT_MMBE_0700_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSTARTXT,"Stock Type",DT_MMBE_0700_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSTARTXT,"DT_MMBE_0700_GET_TEXT_OF_GRIDCELL_0_BSTNDTXT_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButton("Continue   \(Enter\)",False)
'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

