
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


gstrTestCaseName = "Test_P2P - Continuous  positive (Delta posting)"
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

Call SetTextbox("Article","MATNR-LOW","",DT_MB51_1000_ARTICLE,False)
Call SetTextbox("Site","WERKS-LOW","",DT_MB51_1000_SITE,False)
Call TakeScreenShot
Call SetTextbox("Storage Location","LGORT-LOW","",DT_MB51_1000_STORAGE_LOCATION,False)
Call SetTextbox("Posting Date","BUDAT-LOW","",ConvertDate(DT_MB51_1000_POSTING_DATE),False)
Call TakeScreenShot
Call SetVerticalScrollBar(350,False)
Call SetTextbox("Reference","XBLNR-LOW","",DT_MB51_1000_REFERENCE,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

' SelectColumnGuiGrid(gridTitle, gridIndex, columnName, blnIsItPopup)
Call SelectColumnGuiGrid("","","Article Document",False)
Call ClickButton("Sort in Descending Order   \(Ctrl\+Shift\+F4\)",False)

Call ClickButton("Current Display Variant   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_MB51_0841_SEARCH_TERM,True)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)

Call ClickButton("Current Display Variant   \(Ctrl\+F8\)",False)
Call TakeScreenShot
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_MB51_0841_SEARCH_TERM_OCC1,True)
Call TakeScreenShot
' SetComboByKey(attachedTextOrComboName, keyValue)
Call SetComboByKey("Search Direction",0)
Call PressEnter()
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)

Call ClickButton("Current Display Variant   \(Ctrl\+F8\)",False)
Call TakeScreenShot
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_MB51_0841_SEARCH_TERM_OCC2,True)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)

''''' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContent("",1,"Reference","",DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_XBLNR)
Call VerifyGridCellContent("",1,"Posting Date","",ConvertDate(DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT))
Call VerifyGridCellContent("",1,"Movement type","",DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWART)
Call VerifyGridCellContent("",1,"Document Header Text","",DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BKTXT)
Call VerifyGridCellContent("",1,"Reason for Movement","",DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GRUND)

''''GetGridContent(gridTitle, gridIndex, columnName, rowNumber, refColumn, refFieldVal, dataTableColumnName)
''''Call GetGridContent("","","Article Document",1,"","","DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MBLNR")

'''' GetGridContentByRefColumn(gridTitle, gridIndex, refColumn, refFieldVal, columnName, dataTableColumnName)
Call GetGridContentByRefColumn("","","Reference",DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_XBLNR,"Article Document","DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MBLNR_Output")
Call LogOff()
Call FinalStatus ()
'


'Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'Call TakeScreenShot
'Call SetTextbox("Site","IKPF-WERKS","",DT_MI10_0700_SITE,False)    
'Call SetTextbox("Storage Location","IKPF-LGORT","",DT_MI10_0700_STORAGE_LOCATION,false)
''Call SetTextbox("Storage Location","IKPF-LGORT","","0001",false) 
'Call PressEnter()
'Call TakeScreenShot()
''DT_MI10_0731_ARTICLE_DESCRIPTION changed to DTARTICLE
'Call SetTextbox("Material Description","ISEG-MATNR",0,DTARTICLE,false)  
''DT_MI10_0731_ISEGERFMG
'call SetTextboxNoLabel("ISEG-ERFMG",0,DT_MI10_0731_ISEGERFMG,false)  'quantity coumn
'call SetTextboxNoLabel("ISEG-BSTAR",0,DT_MI10_0731_ISEGBSTAR,false)  'sty column
'Call TakeScreenShot()
'Call SetTextbox("Material Description","ISEG-MATNR",1,DTARTICLE_OCC,false)  
'call SetTextboxNoLabel("ISEG-ERFMG",1,DT_MI10_0731_ISEGERFMG_OCC1,false)  'quantity coumn
'call SetTextboxNoLabel("ISEG-BSTAR",1,DT_MI10_0731_ISEGBSTAR_OCC1,false)  'sty column
'Call TakeScreenShot
'Call PressEnter()
'Call TakeScreenShot
'Call ClickButton("btn\[11\]",false)
'
'
'
'
'
'
'
'
'call GetStatusBar("item1","DT_Invoicedoc_Output")
''Call VerifyStatusBar(DT_Invoicedoc_Output)
'
'
'Call LogOff()
'Call FinalStatus ()
'

