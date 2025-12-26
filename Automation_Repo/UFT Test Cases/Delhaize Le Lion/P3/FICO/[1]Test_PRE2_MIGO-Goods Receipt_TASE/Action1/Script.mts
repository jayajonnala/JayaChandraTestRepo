'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_Customer Test_PRE2_MIGO-Goods Receipt_TASE
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE2_MIGO-Goods Receipt_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DLL\FICO\TASE_DT_PRE2_MIGO-Goods Receipt.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''


SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-ZFIAR_AFF_CLEARING----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()

'''Added set combo by key to avoid purchase order selection issues
Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPRO_ACTION_KEY)
Wait 5
Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPRO_REFDOC_KEY)
Wait 5

Call SetTextbox("GR goods receipt","GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULT_TVBWART,False)

Call SetTextBoxNolabel("GODYNPRO-PO_NUMBER",0,DT_MIGO_2000_GODYNPROPO_NUMBER,False)
Call PressEnter()     
Call TakeScreenShot()
Call SetTextBox("Delivery Note","GOHEAD-LFSNR",0,DT_MIGO_0110_DELIVERY_NOTE,False)
Call SelectCheckBox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call ClickButton("btn\[11\]",False)     
Call TakeScreenshot()
Call GetStatusBar("item1","DT_ART_DOC_OUTPUT")
Call LogOff()
Call FinalStatus ()



'*********************************************End Of Script*********************************************************************
