'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_BR0212-MD-03 Luxembourg - simple dry article created via ZWRFMATCOPY and copied from a local Lux article
'.................Author : TCS_Ramesh
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



gstrTestCaseName = "Test_BR0212-MD-03 Luxembourg -  article  Lux _TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\rtod\Documents\Input Datasheet\DLL\DT_04.04.02.12 VIM - NPO Precontrole Issue - BR10b - Invalid Vendor_TASE2.xls"



Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'DataRowSet =2

'SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
Wait(2)

'----------------------Tcode ZMDAM_WRFMATCOPY----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call TakeScreenShot()
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()
 
Call ClickButton("Load selection screen variant   \(Shift\+F5\)",fALSE)
Call SetTextbox("Screen Variant","ZMDAM_MATCOPY_SV-MATCOPY_VARIANT","",DT_ZMDAM_WRFMATCOPY_0201_SCREEN_VARIANT,True)
Call TakeScreenShot()
 
 
 
Call ClickButton("OK   \(F2\)",fALSE)  
Call SetTextbox("Article","ZMDAM_S_MATCOPY_SSCR-P_RMATN","",DT_ZMDAM_WRFMATCOPY_0050_ARTICLE,False) 
Call PressEnter()
Call TakeScreenShot() 
 

Call ClickButton("Key Filters",fALSE) 
Call SelectRadioButton("SPOPLI-SELFLAG","Hierarchy ID",True)
Call TakeScreenShot() 
Call ClickButton("Continue   \(Enter\)",True)

Call ClickButton("Deselect all   \(F6\)",True)
Call SelectRowGuiGridbyRowNo("Reduction Level",0,3,True)
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Wait(2)
Call ClickButton("btn\[8\]",False)
Call TakeScreenShot() 

''Call SetGridData("Variant Filters",1,"GTIN",DT_ZMDAM_WRFMATCOPY_0700_GRIDCELL_0_GTIN,True)
''Call SetGridData("Variant Filters",2,"GTIN",DT_ZMDAM_WRFMATCOPY_0700_GRIDCELL_1_GTIN,True)
''Call SetGridData("Variant Filters",3,"GTIN",DT_ZMDAM_WRFMATCOPY_0700_GRIDCELL_2_GTIN,True)
''Call TakeScreenShot()
''
''Call SelectCheckBoxGridByRowNo("Variant Filters","Main GTIN of vendor",1,0,"ON")
''Call SelectCheckBoxGridByRowNo("Variant Filters","Main GTIN of vendor",2,0,"ON")
''Call SelectCheckBoxGridByRowNo("Variant Filters","Main GTIN of vendor",3,0,"ON")
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(20)
Call TakeScreenShot()

Call GetGridContent("",0,"MESSAGE_V1",1,"Message number","23","DT_ZMDAM_WRFMATCOPY_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MESSAGE_V1_Output")
Call VerifyGridCellContent("",3,"MESSAGE",0,DT_ZMDAM_WRFMATCOPY_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MESSAGE)



'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()



