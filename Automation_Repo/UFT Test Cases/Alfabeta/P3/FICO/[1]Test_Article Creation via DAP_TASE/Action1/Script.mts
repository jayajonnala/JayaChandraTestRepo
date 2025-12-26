
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Article Creation via DAP
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Article Creation via DAP
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 24th Oct
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Article Creation via DAP"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''Close All Browser
Call CloseAllBrowsers()
'
''launch adn Login DAP Application
Call  LaunchSAPWebApplicationEdge(DT_SAPURL)
Wait(10)

Browser("CreationTime:=0").FullScreen

Call ClickWebElement(0, "", "I", "", "caret left-text", 0, False)
CAll ClickLink(0, "", "", "English", False)
Call CaptureWebScreen(0,"Create New Article")


Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


'Click on Create new
Call ClickWebButton(0,"",".*btn-primary btn-block.*|.*btn-default btn-block.*","Create new","BUTTON",0,False)
Call CaptureWebScreen(0,"Create New Article")

''Enter the Basic Details
Call SelectWebList(0,"","Article_View\.Artc_Basic_Info\.country_code_iso",0,DT_ARTICLE_VIEWARTC_BASIC_INFOCOUNTRY_CODE_ISO,False)
Wait(5)
Call SelectWebList(0,"","Article_View\.Artc_Basic_Info\.hdr_matl_type",0,DT_ARTICLE_VIEWARTC_BASIC_INFOVEND_MATL_TYPE,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Basic_Info\.vend_countryori",0,DT_ARTICLE_VIEWARTC_BASIC_INFOVEND_COUNTRYORI,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Enter Basic Details")

Call ClickWebButton(0,"","btn btn-primary btn-lg","","BUTTON",0,False)
Call CaptureWebScreen(0,"Create New Article:Enter Basic Details")
Call SetWebEdit(0, "", "article-brand-modal-search-input", "text", 0, DT_ARTICLE_BRAND_MODAL_EARCH_INPUT, False)
Call ClickWebButton(0,"","btn btn-default","","BUTTON",0,False)
Call DoubleClickWebElement(0, "", "TD", DT_NO_BRAND, "", 0, False)

Call CaptureWebScreen(0,"Create New Article:Enter Basic Details")

Call SelectWebList(0,"","Article_View\.Artc_Basic_Info\.basic_unit_iso",0,DT_ARTICLE_VIEWARTC_BASIC_INFOBASIC_UNIT_ISO,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Basic_Info\.vend_gross_cont","text",0,DT_ARTICLE_VIEWARTC_BASIC_INFOVEND_GROSS_CONT,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Basic_Info\.vend_net_cont","text",0,DT_ARTICLE_VIEWARTC_BASIC_INFOVEND_NET_CONT,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Basic_Info\.vend_cont_unit_iso",0,DT_ARTICLE_VIEWARTC_BASIC_INFOVEND_CONT_UNIT_ISO,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Enter Basic Details")

Call SetWebEdit(0,"","Article_View\.Artc_Basic_Info\.vend_shelf_life","text",0,DT_ARTICLE_VIEWARTC_BASIC_INFOVEND_SHELF_LIFE,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Basic_Info\.vend_rem_shelf_life","text",0,DT_ARTICLE_VIEWARTC_BASIC_INFOVEND_REM_SHELF_LIFE,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Basic_Info\.vend_stor_percent","text",0,DT_ARTICLE_VIEWARTC_BASIC_INFOVEND_STOR_PERCENT,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Basic_Info\.vend_temp_conds",0,"21 - Dry 21°",False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Enter Basic Details")

Call SetWebEdit(0,"","Article_View\.Artc_Basic_Info\.sap_ref_article","text",0,DT_ARTICLE_VIEWARTC_BASIC_INFOSAP_REF_ARTICLE,False)
Wait(1)

'Click on Article description
Call ClickLink(0,"","","Article Descriptions",False)
Wait(8)

'Enter the Article description
Call SelectWebList(0,"","Article_View\.Artc_Descriptions\[0\]\.langu_iso",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS0LANGU_ISO,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[0\]\.matl_vdescr","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS0MATL_VDESCR,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[0\]\.matl_sdescr","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS0MATL_SDESCR,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[0\]\.matl_sldescr1","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS0MATL_SLDESCR1,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[0\]\.matl_tdescr","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS0MATL_SDESCR_OCC1,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Enter Article description")

'Click on Link "Click here to insert another Description Section"
Call ClickWebButton(0,"",".*btn-link btn-xs.*","Click here to insert another Description Section!","BUTTON",0,False)
Wait(5)

Call SelectWebList(0,"","Article_View\.Artc_Descriptions\[1\]\.langu_iso",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS1LANGU_ISO,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[1\]\.matl_vdescr","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS1MATL_VDESCR,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[1\]\.matl_sdescr","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS1MATL_SDESCR,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[1\]\.matl_sldescr1","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS1MATL_SLDESCR1,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Descriptions\[1\]\.matl_tdescr","text",0,DT_ARTICLE_VIEWARTC_DESCRIPTIONS1MATL_TDESCR,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Enter Article description")

Call ClickLink(0,"","","Item notes / Special Instructions",False)
Wait(8)
Call SetWebEdit(0,"","Article_View\.Artc_Special_Notes\[0\]\.item_notes","textarea",0,DT_ARTICLE_VIEWARTC_SPECIAL_NOTES0ITEM_NOTES,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Enter Article description")

'Click on Units of Measure
Call ClickLink(0,"","","Units of Measure",False)
Wait(1)

'Enter the details
Call SelectWebCheckbox(0,"","Article_View\.Artc_Base_Units\[0\]\.vend_deli_uom_flag","INPUT",0,"OFF",False)
Call SelectWebCheckbox(0,"","Article_View\.Artc_Base_Vend_GTINs\[0\]\.vend_gtin_main_flag","INPUT",0,"ON",False)
Call SelectWebCheckbox(0,"","Article_View\.Artc_Base_Units\[0\]\.vend_sales_uom_flag","INPUT",0,"ON",False)
Call SelectWebList(0,"","Article_View\.Artc_Base_Vend_GTINs\[0\]\.vend_gtin_cat",0,DT_ARTICLE_VIEWARTC_BASE_VEND_GTINS0VEND_GTIN_CAT,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Base_Vend_GTINs\[0\]\.vend_gtin","text",0,DT_ARTICLE_VIEWARTC_BASE_VEND_GTINS0VEND_GTIN,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Base_Units\[0\]\.vend_net_weight","text",0,DT_ARTICLE_VIEWARTC_BASE_UNITS0VEND_NET_WEIGHT,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Base_Units\[0\]\.vend_gross_wt","text",0,DT_ARTICLE_VIEWARTC_BASE_UNITS0VEND_GROSS_WT,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Base_Units\[0\]\.vend_length","text",0,DT_ARTICLE_VIEWARTC_BASE_UNITS0VEND_LENGTH,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Base_Units\[0\]\.vend_width","text",0,DT_ARTICLE_VIEWARTC_BASE_UNITS0VEND_WIDTH,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Base_Units\[0\]\.vend_height","text",0,DT_ARTICLE_VIEWARTC_BASE_UNITS0VEND_HEIGHT,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Enter Measure Details")

'Click on Link Alternative UoMs 
Call ClickLink(0,"","","Alternative UoMs",False)
Wait(1)

Call ClickWebButton(0,"",".*btn-link btn-xs.*","Click here to add a new Alternative UOM!","BUTTON",0,False)
Wait(5)
Call SelectWebCheckbox(0,"","Article_View\.Artc_Alt_Units\[1\]\.vend_po_uom_flag","INPUT",0,DT_ARTICLE_VIEWARTC_ALT_UNITS1VEND_PO_UOM_FLAG,False)
Call SelectWebCheckbox(0,"","Article_View\.Artc_Alt_Units\[1\]\.vend_deli_uom_flag","INPUT",0,DT_ARTICLE_VIEWARTC_ALT_UNITS1VEND_DELI_UOM_FLAG,False)
Call SelectWebCheckbox(0,"","Article_View\.Artc_Alt_Vend_GTINs\[1\]\.vend_gtin_main_flag","INPUT",0,"ON",False)
Call SelectWebList(0,"","Article_View\.Artc_Alt_Units\[1\]\.unit_iso",0,DT_ARTICLE_VIEWARTC_ALT_UNITS1UNIT_ISO,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Alt_Units\[1\]\.vend_no_luom","text",0,DT_ARTICLE_VIEWARTC_ALT_UNITS1VEND_NO_LUOM,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Alt_Units\[1\]\.vend_luom",0,DT_ARTICLE_VIEWARTC_ALT_UNITS1VEND_LUOM,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Alt_Vend_GTINs\[1\]\.vend_gtin_cat",0,DT_ARTICLE_VIEWARTC_ALT_VEND_GTINS1VEND_GTIN_CAT,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Add Alternate UOM")

Call ClickWebButton(0,"",".*btn-link btn-xs.*","Click here to add a new Alternative UOM!","BUTTON",0,False)
Call SelectWebCheckbox(0,"","Article_View\.Artc_Alt_Vend_GTINs\[2\]\.vend_gtin_main_flag","INPUT",0,"ON",False)
Call SetWebEdit(0,"","Article_View\.Artc_Alt_Units\[1\]\.vend_gross_wt","text",0,DT_ARTICLE_VIEWARTC_ALT_UNITS1VEND_GROSS_WT,False)
Call SetWebEdit(0,"","Article_View\.Artc_Alt_Units\[1\]\.vend_length","text",0,DT_ARTICLE_VIEWARTC_ALT_UNITS1VEND_LENGTH,False)
Call SetWebEdit(0,"","Article_View\.Artc_Alt_Units\[1\]\.vend_width","text",0,DT_ARTICLE_VIEWARTC_ALT_UNITS1VEND_WIDTH,False)
Call SetWebEdit(0,"","Article_View\.Artc_Alt_Units\[1\]\.vend_height","text",0,DT_ARTICLE_VIEWARTC_ALT_UNITS1VEND_HEIGHT,False)

Call SelectWebList(0,"","Article_View\.Artc_Alt_Units\[2\]\.unit_iso",0,DT_ARTICLE_VIEWARTC_ALT_UNITS2UNIT_ISO,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Alt_Units\[2\]\.vend_no_luom","text",0,DT_ARTICLE_VIEWARTC_ALT_UNITS2VEND_NO_LUOM,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Alt_Units\[2\]\.vend_luom",0,DT_ARTICLE_VIEWARTC_ALT_UNITS2VEND_LUOM,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Alt_Vend_GTINs\[2\]\.vend_gtin_cat",0,DT_ARTICLE_VIEWARTC_ALT_VEND_GTINS2VEND_GTIN_CAT,False)
Wait(1)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Add Alternate UOM")

Call ClickWebButton(0,"",".*btn-link btn-xs.*","Click here to add a new Alternative UOM!","BUTTON",0,False)
Wait(5)
Call SelectWebCheckbox(0,"","Article_View\.Artc_Alt_Vend_GTINs\[3\]\.vend_gtin_main_flag","INPUT",0,"ON",False)
Call SetWebEdit(0,"","Article_View\.Artc_Alt_Units\[2\]\.vend_gross_wt","text",0,DT_ARTICLE_VIEWARTC_ALT_UNITS2VEND_GROSS_WT,False)
Call SetWebEdit(0,"","Article_View\.Artc_Alt_Units\[2\]\.vend_length","text",0,DT_ARTICLE_VIEWARTC_ALT_UNITS2VEND_LENGTH,False)
Call SetWebEdit(0,"","Article_View\.Artc_Alt_Units\[2\]\.vend_width","text",0,DT_ARTICLE_VIEWARTC_ALT_UNITS2VEND_WIDTH,False)
Call SetWebEdit(0,"","Article_View\.Artc_Alt_Units\[2\]\.vend_height","text",0,DT_ARTICLE_VIEWARTC_ALT_UNITS2VEND_HEIGHT,False)


Call SelectWebList(0,"","Article_View\.Artc_Alt_Units\[3\]\.unit_iso",0,DT_ARTICLE_VIEWARTC_ALT_UNITS3UNIT_ISO,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Alt_Units\[3\]\.vend_no_luom","text",0,DT_ARTICLE_VIEWARTC_ALT_UNITS3VEND_NO_LUOM,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Alt_Units\[3\]\.vend_luom",0,DT_ARTICLE_VIEWARTC_ALT_UNITS3VEND_LUOM,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Alt_Vend_GTINs\[3\]\.vend_gtin_cat",0,DT_ARTICLE_VIEWARTC_ALT_VEND_GTINS3VEND_GTIN_CAT,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Add Alternate UOM")

Call ClickWebButton(0,"",".*btn-link btn-xs.*","Click here to add a new Alternative UOM!","BUTTON",0,False)
Wait(8)
Call SelectWebCheckbox(0,"","Article_View\.Artc_Alt_Vend_GTINs\[4\]\.vend_gtin_main_flag","INPUT",0,"ON",False)
Call SelectWebList(0,"","Article_View\.Artc_Alt_Units\[4\]\.unit_iso",0,DT_ARTICLE_VIEWARTC_ALT_UNITS4UNIT_ISO,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Alt_Units\[4\]\.vend_no_luom","text",0,DT_ARTICLE_VIEWARTC_ALT_UNITS4VEND_NO_LUOM,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Alt_Units\[4\]\.vend_luom",0,DT_ARTICLE_VIEWARTC_ALT_UNITS4VEND_LUOM,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Alt_Vend_GTINs\[4\]\.vend_gtin_cat",0,DT_ARTICLE_VIEWARTC_ALT_VEND_GTINS4VEND_GTIN_CAT,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Add Alternate UOM")

'Click on Link Purchasing 
Call ClickLink(0,"","","Purchasing",False)
Wait(1)

'Enter the purchasing details
Call SetWebEdit(0,"","Article_View\.Artc_Vendor_Data\.vend_lifnr","text",0,DT_ARTICLE_VIEWARTC_VENDOR_DATAVEND_LIFNR,False)
Wait(8)
Call SetWebEdit(0,"","Article_View\.Artc_Vendor_Data\.hdr_vend_mat","text",0,DT_ARTICLE_VIEWARTC_VENDOR_DATAHDR_VEND_MAT,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Enter Purchasing  Details")

'Click on Purchasing Group Data
Call ClickLink(0,"","","Purchasing Group Data",False)
Wait(1)

'Enter the Purchasing Group Data Details
Call SelectWebList(0,"","Article_View\.Artc_Vend_Tax_Datas\[0\]\.vend_tax_type",0,"VAT - VAT tax",False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Vend_Tax_Datas\[0\]\.vend_tax_rate",0,DT_ARTICLE_VIEWARTC_VEND_TAX_DATAS0VEND_TAX_RATE,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Vend_Pricing_Infos\[0\]\.purch_ekgrp",0,DT_ARTICLE_VIEWARTC_VEND_PRICING_INFOS0PURCH_EKGRP,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Vend_Pricing_Infos\[0\]\.vend_netpr","text",0,DT_ARTICLE_VIEWARTC_VEND_PRICING_INFOS0VEND_NETPR,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_Vend_Pricing_Infos\[0\]\.vend_waers_iso",0,DT_ARTICLE_VIEWARTC_VEND_PRICING_INFOS0VEND_WAERS_ISO,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_Vend_Pricing_Infos\[0\]\.vend_ord_pr_uom_iso","text",0,DT_ARTICLE_VIEWARTC_VEND_PRICING_INFOS0VEND_ORD_PR_UOM_ISO,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:Enter Purchasing Group Data Details")


'Click here to add another Condition type
Call ClickWebButton(0,"",".*btn-link btn-xs.*","Click here to add another Condition type!","BUTTON",0,False)
Wait(10)

Call SelectWebList(0,"","Article_View\.Artc_Vend_Other_Conds\[0\]\.cond_type",0,DT_ARTICLE_VIEWARTC_VEND_OTHER_CONDS0COND_TYPE,False)
Wait(8)

Browser("Delhaize Article Portal").Page("Delhaize Article Portal").Link("Logistics").Click @@ script infofile_;_ZIP::ssf1.xml_;_
' SelectWebList(creationTime, windowText, weblistName, weblistIndex, val, blnIsItPopup)
Call CaptureWebScreen(0,"Create New Article:OpCo Specific View")
Call SelectWebList(0,"","Article_View\.Artc_MDM_ABs\[0\]\.logistics_dsd_central",0,DT_DSD_OR_CENTRAL,False)
Wait(10)
Call ClickWebButton(0,"",".*primary save-button.*","Save","BUTTON",0,False)
Wait(5)
Call CaptureWebScreen(0,"Create New Article:Save Details")
Call ClickWebButton(0, "", "btn dropdown-toggle selectpicker btn-default", " Select\.\.\. ", "BUTTON", 0, False)
Wait 1
Call ClickWebElement(0, "", "SPAN",DT_WAREHOUSE, "text", 0, False)
Wait 5
Call SelectWebList(0,"","Article_View\.Artc_MDM_ABs\[0\]\.logistics_itemset",0,DT_LOGISTICS_ITEMSET,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_MDM_ABs\[0\]\.logistics_max_life_in_store","text",0,DT_logistics_max_life_in_store,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_MDM_ABs\[0\]\.logistics_pricing_policy",0,DT_logistics_pricing_policy,False)
Wait(1)
Call CaptureWebScreen(0,"Create New Article:OpCo Specific View")

'Click on OpCo Specific View
Call ClickLink(0,"","","OpCo Specific View",False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_MDM_ABs\[0\]\.self_service",0,DT_ARTICLE_VIEWARTC_MDM_ABS0SELF_SERVICE,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_MDM_ABs\[0\]\.sesonality_flag",0,DT_ARTICLE_VIEWARTC_MDM_ABS0SESONALITY_FLAG,False)
Wait(1)

Call SelectWebList(0,"","Article_View\.Artc_MDM_ABs\[0\]\.storage_type",0,DT_ARTC_MDM_ABS_STORAGE_TYPE,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_MDM_ABs\[0\]\.min_order_qunatity","text",0,DT_ARTC_MDM_ABS_MIN_ORDER_QUNATITY,False)
Wait(1)
Call SetWebEdit(0,"","Article_View\.Artc_MDM_ABs\[0\]\.max_order_qunatity","text",0,DT_ARTC_MDM_ABS_MAX_ORDER_QUNATITY,False)
Wait(1)
Call SelectWebList(0,"","Article_View\.Artc_MDM_ABs\[0\]\.assessment_unit",0,DT_ARTICLE_VIEW_ARTC_MDM_ABS_ASSESSMENT_UNIT,False)
Wait(1)

Call CaptureWebScreen(0,"Create New Article:OpCo Specific View")



'Click on Save details
Call ClickWebButton(0,"",".*primary save-button.*","Save","BUTTON",0,False)
Wait(5)
Call CaptureWebScreen(0,"Create New Article:Save Details")

'Verify If Success Message is generated.
Call VerifyWebElement(0,"","LI","Success! Article saved successfully!",".*",0,False)

'Click on Submit
Call ClickWebButton(0,"",".*default btn-block.*","Submit","BUTTON",0,False)
Wait(15)
Call CaptureWebScreen(0,"Create New Article:Submit Details")

'Click on OK Button
Call ClickWebButton(0,"",".*btn-primary.*","Submit","BUTTON",1,True)
Wait(30)

''Refresh the Data
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
Call GetValueWebTable(0,"table table-condensed.*",1,2,"DT_MAIN_GTIN")
Call GetValueWebTable(0,"table table-condensed.*",2,2,"DT_DESCRIPTION")

Call VerifyWebTableValue(0,"table table-condensed.*","Main GTIN",5,2,"Submitted")

Browser("CreationTime:=0").Maximize

'------------------------'Log Off From Applicaton--------------------------------

'Close All Browser
Call CloseAllBrowsers()

Call FinalStatus ()

'#####################################################End of Script##############################################################

