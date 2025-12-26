
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.1.1.6. Create Article via Web Application - Common article_P1
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.1.1.6. Create Article via Web Application - Common article_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.1.1.6. Create Article via Web Application - Common article_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

''Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''Close All Browser
Call CloseAllBrowsers()

''launch adn Login DAP Application
Call  LaunchSAPWebApplicationEdge(DT_SAPURL)
Wait(10)

Browser("CreationTime:=0").FullScreen
Wait 5
Call ClickWebElement(0, "", "I", "", "caret left-text", 0, False)

CAll ClickLink(0, "", "", "English", False)
Call CaptureWebScreen(0,"Create New Article")

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Click on Create new
Call ClickWebButton(0,"",".*btn-primary btn-block.*|.*btn-default btn-block.*","Create new","BUTTON",0,False)
Call CaptureWebScreen(0,"Create New Article")

'Enter the Basic Details
Call SelectWebList(0,"","Article_View\.Artc_Basic_Info\.country_code_iso",0,DT_ARTICLE_VIEWARTC_BASIC_INFOCOUNTRY_CODE_ISO,False)
Wait(5)
Call CaptureWebScreen(0,"Create New Article")
Call SelectWebList(0,"","Article_View\.Artc_Basic_Info\.hdr_matl_type",0,DT_ARTICLE_VIEWARTC_BASIC_INFOHDR_MATL_TYPE,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Basic_Info\.hdr_matl_cat",0,DT_ARTICLE_VIEWARTC_BASIC_INFOHDR_MATL_CAT,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Basic_Info\.vend_countryori",0,DT_ARTICLE_VIEWARTC_BASIC_INFOVEND_COUNTRYORI,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article")
Call SetWebEdit(0,"","Article_View\.Artc_Basic_Info\.vend_shelf_life","text",0,DT_ARTICLE_VIEWARTC_BASIC_INFOVEND_SHELF_LIFE,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Basic_Info\.vend_rem_shelf_life","text",0,DT_ARTICLE_VIEWARTC_BASIC_INFOVEND_REM_SHELF_LIFE,False)
Wait(1)
'Call SetWebEdit(0,"","Article_View\.Artc_Basic_Info\.vend_stor_percent","text",0,DT_ARTICLE_VIEWARTC_BASIC_INFOVEND_STOR_PERCENT,False)
'Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Basic_Info\.vend_temp_conds",0,"02 - Fresh 2°",False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article")
Call SelectWebList(0,"","Article_View\.Artc_Basic_Info\.basic_unit_iso",0,DT_ARTICLE_VIEWARTC_BASIC_INFOBASIC_UNIT_ISO,False)
Wait(8)
Call SetWebEdit(0,"","Article_View\.Artc_Basic_Info\.vend_net_cont","text",0,DT_ARTICLE_VIEWARTC_BASIC_INFOVEND_GROSS_CONT,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Basic_Info\.vend_gross_cont","text",0,DT_ARTICLE_VIEWARTC_BASIC_INFOVEND_GROSS_CONT,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article")
Wait(1)

Call SelectWebList(0,"","Article_View\.Artc_Basic_Info\.vend_cont_unit_iso",0,DT_ARTICLE_VIEWARTC_BASIC_INFOVEND_CONT_UNIT_ISO ,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Enter Basic Details")

'''Click on Article description
Call ClickLink(0,"","","Article Descriptions",False)

'''Enter the Article description
Call SelectWebList(0,"","Article_View\.Artc_Descriptions\[0\]\.langu_iso",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS0LANGU_ISO,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[0\]\.matl_vdescr","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS0MATL_VDESCR,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[0\]\.vend_func_name","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS0VEND_FUNC_NAME,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[0\]\.vend_variant","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS0VEND_VARIANT,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[0\]\.matl_ldescr","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS0MATL_LDESCR,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[0\]\.matl_sdescr","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS0MATL_SDESCR,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[0\]\.matl_tdescr","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS0MATL_TDESCR,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Enter Article description")

''''Click on Link "Click here to insert another Description Section"
Call ClickWebButton(0,"",".*btn-link btn-xs.*","Click here to insert another Description Section!","BUTTON",0,False)
Wait(5)

''''Enter the Article description
Call SelectWebList(0,"","Article_View\.Artc_Descriptions\[1\]\.langu_iso",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS1LANGU_ISO,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[1\]\.matl_vdescr","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS1MATL_VDESCR,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[1\]\.vend_func_name","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS1VEND_FUNC_NAME,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[1\]\.vend_variant","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS1VEND_VARIANT,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[1\]\.matl_ldescr","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS1MATL_LDESCR,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[1\]\.matl_sdescr","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS1MATL_SDESCR,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[1\]\.matl_tdescr","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS1MATL_TDESCR,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Enter Article description")


'''Click on Units of Measure
Call ClickLink(0,"","","Units of Measure",False)
Wait(1)

'''Enter the details

Call SelectWebCheckbox(0,"","Article_View\.Artc_Base_Vend_GTINs\[0\]\.vend_gtin_main_flag","INPUT",0,"ON",False)
Call SelectWebCheckbox(0,"","Article_View\.Artc_Base_Units\[0\]\.vend_po_uom_flag","INPUT",0,"ON",False)
Call SelectWebCheckbox(0,"","Article_View\.Artc_Base_Units\[0\]\.vend_deli_uom_flag","INPUT",0,"ON",False)
Call SelectWebCheckbox(0,"","Article_View\.Artc_Base_Units\[0\]\.vend_sales_uom_flag","INPUT",0,"ON",False)


Call SelectWebList(0,"","Article_View\.Artc_Base_Vend_GTINs\[0\]\.vend_gtin_cat",0,DT_ARTICLE_VIEWARTC_BASE_VEND_GTINS0VEND_GTIN_CAT,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Base_Vend_GTINs\[0\]\.vend_gtin","text",0,DT_ARTICLE_VIEWARTC_BASE_VEND_GTINS0VEND_GTIN,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Base_Units\[0\]\.vend_net_weight","text",0,DT_ARTICLE_VIEWARTC_BASE_UNITS0VEND_NET_WEIGHT,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Base_Units\[0\]\.vend_gross_wt","text",0,DT_ARTICLE_VIEWARTC_BASE_UNITS0VEND_GROSS_WT,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Base_Units\[0\]\.vend_unit_of_wt",0,DT_ARTICLE_VIEWARTC_BASE_UNITS0VEND_UNIT_OF_WT,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Base_Units\[0\]\.vend_length","text",0,DT_ARTICLE_VIEWARTC_BASE_UNITS0VEND_LENGTH,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Base_Units\[0\]\.vend_width","text",0,DT_ARTICLE_VIEWARTC_BASE_UNITS0VEND_WIDTH,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Base_Units\[0\]\.vend_height","text",0,DT_ARTICLE_VIEWARTC_BASE_UNITS0VEND_HEIGHT,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Base_Units\[0\]\.vend_unit_dim",0,DT_ARTICLE_VIEWARTC_BASE_UNITS0VEND_UNIT_DIM,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Enter Measure Details")

'''Click on Link Purchasing 
Call ClickLink(0,"","","Purchasing",False)
Wait(1)

'''Enter the purchasing details
Call SetWebEdit(0,"","Article_View\.Artc_Vendor_Data\.vend_lifnr","text",0,DT_ARTICLE_VIEWARTC_VENDOR_DATAVEND_LIFNR,False)
Wait(8)
Call SetWebEdit(0,"","Article_View\.Artc_Vendor_Data\.hdr_vend_mat","text",0,DT_ARTICLE_VIEWARTC_VENDOR_DATAHDR_VEND_MAT,False)
Wait(1)
'Call SelectWebList(0,"","Article_View\.Artc_Vendor_Data\.vend_ltsnr",0,DT_ARTICLE_VIEWARTC_VENDOR_DATAVEND_LTSNR,False)
Wait(1)
Call SelectWebCheckbox(0,"","Article_View\.Artc_Vendor_Data\.vend_relif","INPUT",0,"ON",False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Enter Purchasing  Details")

'''Click on Purchasing Group Data
Call ClickLink(0,"","","Purchasing Group Data",False)
Wait(1)

''''Enter the Purchasing Group Data Details
Call SelectWebList(0,"","Article_View\.Artc_Vend_Tax_Datas\[0\]\.vend_tax_type",0,"VAT - VAT tax",False)
Wait(1)
''Call SelectWebList(0,"","Article_View\.Artc_Vend_Tax_Datas\[0\]\.vend_tax_rate",0,DT_ARTICLE_VIEWARTC_VEND_TAX_DATAS0VEND_TAX_RATE,False)
Call SelectWebList(0,"","Article_View\.Artc_Vend_Tax_Datas\[0\]\.vend_tax_rate",0,"No tax",False)                                     
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Vend_Pricing_Infos\[0\]\.purch_ekgrp",0,DT_ARTICLE_VIEWARTC_VEND_PRICING_INFOS0PURCH_EKGRP,False)
Wait(1)
'Call SetWebEdit(0,"","Article_View\.Artc_Vend_Pricing_Infos\[0\]\.vend_netpr","text",0,DT_ARTICLE_VIEWARTC_VEND_PRICING_INFOS0VEND_NETPR,False)
Call SetWebEdit(0,"","Article_View\.Artc_Vend_Pricing_Infos\[0\]\.vend_netpr","text",0,"50",False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Vend_Pricing_Infos\[0\]\.vend_waers_iso",0,DT_ARTICLE_VIEWARTC_VEND_PRICING_INFOS0VEND_WAERS_ISO,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Vend_Pricing_Infos\[0\]\.vend_ord_pr_uom_iso","text",0,DT_ARTICLE_VIEWARTC_VEND_PRICING_INFOS0VEND_ORD_PR_UOM_ISO,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Enter Purchasing Group Data Details")


'''Click on Save details
Call ClickWebButton(0,"",".*primary save-button.*","Save","BUTTON",0,False)
Wait(5)
Call CaptureWebScreen(0,"Create New Article:Save Details")

'''Verify If Success Message is generated.
Call VerifyWebElement(0,"","LI","Success! Article saved successfully!",".*",0,False)

'''Click on Submit
Call ClickWebButton(0,"",".*default btn-block.*","Submit","BUTTON",0,False)
Wait(15)

Call CaptureWebScreen(0,"Create New Article:Submit Details")

'''Click on Submit Button
Call ClickWebButton(0,"",".*btn-primary.*","Submit","BUTTON",1,True)
Wait(30)


'Refresh the Data
'Set Obj=Browser("opentitle:=Delhaize Article Portal.*","title:=Delhaize Article Portal.*").Page("title:=Delhaize Article Portal.*")
'If Obj.WebElement("class:=.*fa-refresh.*","html tag:=I","Index:=0").Exist(5) Then
'	Obj.WebElement("class:=.*fa-refresh.*","html tag:=I","Index:=0").Highlight
'	Obj.WebElement("class:=.*fa-refresh.*","html tag:=I","Index:=0").Click
'	Wait(5)
'End If

'Click on Submitted Link
Call ClickLink(0,"","","Submitted",False)
Wait(8)
Call CaptureWebScreen(0,"Create New Article:Verify Status")


Call GetValueWebTable(0,"table table-condensed.*",5,2,"DT_STATUS")

Call VerifyWebTableValue(0,"table table-condensed.*","Main GTIN",5,2,"Submitted")

'------------------------'Log Off From Applicaton--------------------------------

'Close All Browser
Call CloseAllBrowsers()

Call FinalStatus ()

'#####################################################End of Script##############################################################

