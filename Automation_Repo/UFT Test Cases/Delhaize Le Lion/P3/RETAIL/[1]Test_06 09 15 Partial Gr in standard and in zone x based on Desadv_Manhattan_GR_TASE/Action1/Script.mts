
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_136_SRM Regularization SC Generic item on cost Center (MI)
'.................Author : TCS 	   :
'................ Creation Date    :
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


gstrTestCaseName = "Test_06 09 15 Partial Gr Desadv_Manhattan_GR"
'gstrTestCaseName = "Test_Manhattan_GR_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\TASEWork\Data\TASE_DT_114_SRM Standard SC Free Text on Cost Cent P1_.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution1(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------

'
''''''''Close All Browser
Call CloseAllBrowsers()
'Incrementing value in datasheet
Call WriteRunTimeDataToExcelGlobalSheet ("DT_TIME",Cstr(Hour(now())&":"&Minute(Now())))
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+2))

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet) 
'''Launch Manhattan URL
Call LaunchSAPWebApplicationEdge(DT_URL)
wait 20
 'Login 
Call LoginManhattan(0,DT_USER,DT_PASSWORD)
wait 30
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
Call CaptureWebScreenNew(0,"Capture Screen")

'''click hamburger icon
''''Call ClickWebElementNoreportingText(0,"","SPAN","Main Menu","x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-topbar-menu-icon",0,False)
Call ClickWebElementNoreportingText(0,"","SPAN","Main Menu","x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-topbar-menu-icon",0,False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'Click Show All button
Call ClickWebElement(0,"","SPAN","Show All","x-btn-inner x-btn-inner-default-medium",0,False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'Click Ditribution tab
Call ClickWebElement(0,"","SPAN","Distribution","x-tab-inner x-tab-inner-default",0,False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'Click Purchase Orders
Call ClickWebElement(0,"","SPAN","Purchase Orders","",0,False)
wait 5
''PRIMARY FIELD
''''''''''Call SetWebEdit(0, "", "combobox-1602-inputEl", "text", 0, DT_PO_COMBO_KEY, False)
'AIUtil.SetContext Browser("creationtime:=0")
'AIUtil("combobox", "", micFromTop, 1).Type(DT_PO_COMBO_KEY)
'wait 5
Call SetWebEdit(0, "", "combobox.*inputEl", "text", 2, DT_PO_COMBO_KEY, False)
''Enter the PO number
Call SetWebEdit(0, "", "PurchaseOrder", "text", 0, DT_PO, False)
'Click Apply
Call ClickWebElement(0,"","SPAN","Apply","x-btn-inner x-btn-inner-default-small",0,False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'Verify PO status
Call VerifyWebElement(0,"","DIV",DT_PO_STATUS,"x-grid-cell-inner",0,False)
'''Close the screen
'Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)

'''click hamburger icon
Call ClickWebElementNoreportingText(0,"","SPAN","Main Menu","x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-topbar-menu-icon",0,False)
'Click Show All button
Call ClickWebElement(0,"","SPAN","Show All","x-btn-inner x-btn-inner-default-medium",0,False)
'Click Ditribution tab
Call ClickWebElement(0,"","SPAN","Distribution","x-tab-inner x-tab-inner-default",0,False)
'Click Create ASN From PO
Call ClickWebElement(0,"","SPAN","Create ASN From PO","",0,False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'Enter PO nubmer
Call SetWebEdit(0, "", "dataForm:filterId:field10value1", "text", 0, DT_PO, False)
'click Apply
Call ClickWebButtonHtmlId(0, "", "btn", "dataForm:filterId:filterIdapply", "INPUT", 0, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'click Add
Call ClickWebButtonHtmlId(0, "", "btn", "dataForm:cbaddasn", "INPUT", 0, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'click Generate button
Call ClickWebButtonHtmlId(0, "", "", "dataForm:gena", "INPUT", 0, False)
'GET ASN VALUE
Call GetValueWebElementNoInnerText(0, "dataForm:asnidh1","INPUT",0,0,"DT_ASN_OUTPUT")
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
Call SetWebEdit(0, "", "dataForm:sdqtyhcc", "text", 0, DT_DELIV_DATE, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'click Ok
Call ClickWebButtonHtmlId(0, "", "", "dataForm:sv", "INPUT", 0, False)
'select PO checkbox
Call SelectWebCheckbox(0, "", "dataForm:treeTable:0:treeB:adaptor:0::selectId", "INPUT", 0, "ON", False)
'select ASN checkbox
Call SelectWebCheckbox(0, "", "dataForm:atreeTable:0:atreeB:aadaptor:0::aselectId", "INPUT", 0, "ON", False)
'click move button
Call ClickWebElementById(0,"","INPUT","dataForm:cb5","",0,False)
'Expand ASN
Call ClickWebElementById(0,"","IMG","dataForm:atreeTable:0:atreeB:aadaptor:0::j_id168:handle:img:collapsed","rich-tree-node-handleicon-collapsed",0,False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'''Maximize the screen
''Call ClickWebElement(0,"","IMG", "", "x-tool-img x-tool-maximize", 0, False)
'Verify PO 
Call VerifyWebElement(0, "", "SPAN", DT_PO, "h-outputTextD", 0, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
' click Save
Call ClickWebButton(0, "", "btn", "Save", "INPUT", 0, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
'''click hamburger icon
Call ClickWebElementNoreportingText(0,"","SPAN","Main Menu","x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-topbar-menu-icon",0,False)
''Click Show All button
Call ClickWebElement(0,"","SPAN","Show All","x-btn-inner x-btn-inner-default-medium",0,False)
'Click Ditribution tab
Call ClickWebElement(0,"","SPAN","Distribution","x-tab-inner x-tab-inner-default",0,False)
'Click Purchase Orders
Call ClickWebElement(0,"","SPAN","Purchase Orders","",0,False)
wait 5
'''PRIMARY FIELD
Call SetWebEdit(0, "", "combobox.*inputEl", "text", 2, DT_PO_COMBO_KEY, False)

'''AIUtil.SetContext Browser("creationtime:=0")
'''AIUtil("combobox", "", micFromTop, 2).Type(DT_PO_COMBO_KEY)
'Enter the PO number
Call SetWebEdit(0, "", "PurchaseOrder", "text", 0, DT_PO, False)
'Click Apply
Call ClickWebElement(0,"","SPAN","Apply","x-btn-inner x-btn-inner-default-small",0,False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'Verify PO status
Call VerifyWebElement(0,"","DIV",DT_PO_STATUS_1,"x-grid-cell-inner",0,False)
'Select CheckBox
Call ClickWebElement(0, "", "DIV", "", "x-grid-row-checker", 0, False)
''Click View
Call ClickWebElement(0, "", "SPAN", "View", "x-btn-inner x-btn-inner-default-small", 0, False)
'''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
''Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
wait 10
'Click ASN tab
Call ClickLink(0, "", "tab_link", "ASNs", False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Verify ASN number
Call VerifyWebElement(0, "", "SPAN", DT_ASN, "", 0, False)
'''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
'''click hamburger icon
Call ClickWebElementNoreportingText(0,"","SPAN","Main Menu","x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-topbar-menu-icon",0,False)
'Click Show All button
Call ClickWebElement(0,"","SPAN","Show All","x-btn-inner x-btn-inner-default-medium",0,False)
'Click Ditribution tab
Call ClickWebElement(0,"","SPAN","Distribution","x-tab-inner x-tab-inner-default",0,False)
'Click ASNs
Call ClickWebElement(0,"","SPAN","ASNs","",0,False)
WAIT 5
'PRIMARY FIELD
''AIUtil.SetContext Browser("creationtime:=0")
''AIUtil("combobox", "", micFromTop, 1).Type(DT_ASN_COMBO_KEY)
Call SetWebEdit(0, "", "combobox.*inputEl", "text", 2, DT_ASN_COMBO_KEY, False)
'Enter the ASN number
Call SetWebEdit(0, "", "asnId", "text", 0, DT_ASN, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'Click Apply
Call ClickWebElement(0,"","SPAN","Apply","x-btn-inner x-btn-inner-default-small",0,False)
wait 5
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''Select CheckBox
Call ClickWebElement(0, "", "DIV", "", "x-grid-row-checker", 0, False)
''Click Edit Header
Call ClickWebElement(0, "", "SPAN", "Edit Header", "x-btn-inner x-btn-inner-default-small", 0, False)
'wait 5
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
''Enter ARC code
Call SetWebEdit(0, "", "dataForm:Ref_Field2_optxt", "text", 0, DT_ARC_CODE, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'''click Save
Call ClickWebButton(0, "", "btn", "Save", "INPUT", 0, False)
'''''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
'click hamburger icon
Call ClickWebElementNoreportingText(0,"","SPAN","Main Menu","x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-topbar-menu-icon",0,False)
'Click Show All button
Call ClickWebElement(0,"","SPAN","Show All","x-btn-inner x-btn-inner-default-medium",0,False)
'Click Configuration tab
Call ClickWebElementById(0,"", "SPAN","tab-1100-btnInnerEl","x-tab-inner x-tab-inner-default", 2, False)
'Click Items
Call ClickWebElement(0,"","SPAN","Items","",0,False)
wait 5
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'''Enter the article number
Call SetWebEdit(0, "", "dataForm:ItemList_lv:ItemList_filterId:itemLookUpId", "text", 0, DT_ARTICLE, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''click Apply
Call ClickWebButton(0, "", "btn  groupBtn", "Apply", "INPUT", 0, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''select article Checkbox
Call SelectWebCheckbox(0, "", "checkAll_c0_dataForm:ItemList_lv:dataTable", "INPUT", 0, "ON", False)
''click View
Call ClickWebButton(0, "", "btn", "View", "INPUT", 0, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''Verify Teach in Indicator
Call VerifyWebElement(0, "", "SPAN", DT_TEACH_IN_INDICATOR, "captionData", 0, False)
''click Item Facility
Call ClickWebButton(0, "", "btn", "Item Facility", "INPUT", 0, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''select article Checkbox
Call SelectWebCheckbox(0, "", "checkAll_c0_dataForm:ItemFacilityList_lv:dataTable", "INPUT", 0, "ON", False)
''click View
Call ClickWebButton(0, "", "btn", "View", "INPUT", 0, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'GET Putaway type
Call GetWebElementValue(0,"dataForm:ItemFacilityDetails_SOMPutawayType","captionData","SPAN",0,0,"DT_PUTAWAY_TYPE_OUTPUT")
'GET Allocation type
Call GetWebElementValue(0,"dataForm:ItemFacilityDetails_SOMAllocationType","captionData","SPAN",0,0,"DT_ALLOCATION_TYPE_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
''click hamburger icon
Call ClickWebElementNoreportingText(0,"","SPAN","Main Menu","x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-topbar-menu-icon",0,False)
'Click Show All button
Call ClickWebElement(0,"","SPAN","Show All","x-btn-inner x-btn-inner-default-medium",0,False)
'Click Ditribution tab
Call ClickWebElement(0,"","SPAN","Distribution","x-tab-inner x-tab-inner-default",0,False)
'Click Dock Door
Call ClickWebElement(0,"","SPAN","Dock Door","",0,False)
'''dock door workaround
'''intId=0
'''Dim strDock,strStatus,strHtmlId
'''While strDock<>DT_DOCK_DOOR_VALUE or strStatusValue<>DT_STATUS
'''	strHtmlIdDock="dataForm:listView:dataTable:"&Cstr(intId)&":custId3"
'''	strHtmlIdStatus="dataForm:listView:dataTable:"&Cstr(intId)&":out5"
'''	strDock=CheckValueWebElement(0, strHtmlIdDock, DT_DOCK_DOOR_VALUE, "SPAN", 0, 0)
'''	strStatusValue=CheckValueWebElement(0,strHtmlIdStatus,DT_STATUS,"SPAN",0,0)
'''	intId=intId+1
'''Wend
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
wait 5
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'Verify Allocation
Call VerifyWebElementById(0,"","SPAN", DT_DOCK_DOOR,"dataForm:listView:dataTable:"&DT_HTML_ID&":custId3",0,False)
''Verify Status
Call VerifyWebElementById(0,"","SPAN", DT_STATUS,"dataForm:listView:dataTable:"&DT_HTML_ID&":out5",0,False)
'Verify Dock Door
Call VerifyWebElementById(0,"","SPAN", DT_DOCK_DOOR_VALUE,"dataForm:listView:dataTable:"&DT_HTML_ID&":out4",0,False)
''''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
'click hamburger icon
Call ClickWebElementNoreportingText(0,"","SPAN","Main Menu","x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-topbar-menu-icon",0,False)
'Click Show All button
Call ClickWebElement(0,"","SPAN","Show All","x-btn-inner x-btn-inner-default-medium",0,False)
'Click Ditribution tab
Call ClickWebElement(0,"","SPAN","Distribution","x-tab-inner x-tab-inner-default",0,False)
'Click RF Menu
Call ClickWebElement(0,"","SPAN","RF Menu","",0,False)
wait 5
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'click Receiving
Call ClickLink(0, "", "menubutton", "Receiving", False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''''click warehouse type
Call ClickLink(0, "", "menubutton", DT_RF_MENU, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
Call ClickLink(0, "", "menubutton", "BlindRecvMan", False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'''''Enter dock door value
''''''Call SetWebEdit(0, "", "DockInp", "text", 0, DT_DOCK_DOOR_VALUE, False)
Call SetWebEditAndSubmit(0, "", "DockInp", "text", 0, DT_DOCK_DOOR_VALUE, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'Enter ASN value
Call SetWebEdit(0, "", "shipinpId", "text", 0, DT_ASN, False)
Call SetWebEditAndSubmit(0, "", "shipinpId", "text", 0, DT_ASN, False)
''Enter LPN value
lpnValue1=Cstr(DT_ARTICLE)+Cstr(DT_LPN_INC_1)
Call SetWebEditAndSubmit(0,"","lpninput","text",0,lpnValue1,False)
'''Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''Enter Item Barcode
Call SetWebEditAndSubmit(0,"","verfiyItemBrcd","text",0,DT_ARTICLE,False)
Call  SendWebKey("ENTER")
''Enter Qty
Call SetWebEditAndSubmit(0,"","input1input2","text",0,DT_QTY,False)
Call  SendWebKey("ENTER")
wait 5
''Enter LPN value
lpnValue2=Cstr(DT_ARTICLE)+Cstr(DT_LPN_INC_2)
Call SetWebEditAndSubmit(0,"","lpninput","text",0,lpnValue2,False)
'''Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'Enter Item Barcode
Call SetWebEditAndSubmit(0,"","verfiyItemBrcd","text",0,DT_ARTICLE,False)
Call  SendWebKey("ENTER")
'Enter Qty
Call SetWebEditAndSubmit(0,"","input1input2","text",0,DT_QTY,False)
Call  SendWebKey("ENTER")
''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)

''click hamburger icon
Call ClickWebElementNoreportingText(0,"","SPAN","Main Menu","x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-topbar-menu-icon",0,False)
'''Click Show All button
Call ClickWebElement(0,"","SPAN","Show All","x-btn-inner x-btn-inner-default-medium",0,False)
''''Click Ditribution tab
Call ClickWebElement(0,"","SPAN","Distribution","x-tab-inner x-tab-inner-default",0,False)
'Click ASNs
Call ClickWebElement(0,"","SPAN","ASNs","",0,False)
WAIT 5
'PRIMARY FIELD
''AIUtil.SetContext Browser("creationtime:=0")
''AIUtil("combobox", "", micFromTop, 1).Type(DT_ASN_COMBO_KEY)
Call SetWebEdit(0, "", "combobox.*inputEl", "text", 2, DT_ASN_COMBO_KEY, False)
'Enter the ASN number
Call SetWebEdit(0, "", "asnId", "text", 0, DT_ASN, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'Click Apply
Call ClickWebElement(0,"","SPAN","Apply","x-btn-inner x-btn-inner-default-small",0,False)
wait 5
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''Select CheckBox
Call ClickWebElement(0, "", "DIV", "", "x-grid-row-checker", 0, False)
''Click View 
Call ClickWebElement(0, "", "SPAN", "View", "x-btn-inner x-btn-inner-default-small", 0, False)
'''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)

'wait 5
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''click LPN tab
Call ClickLink(0, "", "tab_link", "LPNs", False)
wait 10
'''verify LPNs
Call VerifyWebLinkOutertext(0,"","A",lpnValue1,0,False)
Call VerifyWebLinkOutertext(0,"","A",lpnValue2,0,False)
''Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")

'''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
''click hamburger icon
Call ClickWebElementNoreportingText(0,"","SPAN","Main Menu","x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-topbar-menu-icon",0,False)
'Click Show All button
Call ClickWebElement(0,"","SPAN","Show All","x-btn-inner x-btn-inner-default-medium",0,False)
'Click Ditribution tab
Call ClickWebElement(0,"","SPAN","Distribution","x-tab-inner x-tab-inner-default",0,False)
'Click ASNs
Call ClickWebElement(0,"","SPAN","ASNs","",0,False)
WAIT 5
'PRIMARY FIELD
''AIUtil.SetContext Browser("creationtime:=0")
''AIUtil("combobox", "", micFromTop, 1).Type(DT_ASN_COMBO_KEY)
Call SetWebEdit(0, "", "combobox.*inputEl", "text", 2, DT_ASN_COMBO_KEY, False)
'Enter the ASN number
Call SetWebEdit(0, "", "asnId", "text", 0, DT_ASN, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'Click Apply
Call ClickWebElement(0,"","SPAN","Apply","x-btn-inner x-btn-inner-default-small",0,False)
wait 5
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''Select CheckBox
Call ClickWebElement(0, "", "DIV", "", "x-grid-row-checker", 0, False)
''Click More
Call ClickWebElement(0, "", "SPAN", "More", "x-btn-inner x-btn-inner-default-small", 0, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''Click Verify ASN
Call ClickWebElement(0, "", "SPAN", "Verify ASN", "x-menu-item-text x-menu-item-text-default x-menu-item-indent", 0, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
''click Verify ASN button
Call ClickWebButton(0, "", "btn", "Verify ASN", "INPUT", 0, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''sign out
Call ClickWebElement(0, "", "SPAN", "", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-user ", 0, False)
Call ClickWebElement(0, "", "SPAN","Sign out", "x-btn-inner x-btn-inner-default-small", 0, False)
Call CloseBrowser(0)
Call FinalStatus()


'Public Function ClickWebElementNoreportingText(creationTime,windowText,htmltag,reportingText,className,indexNumber,blnIsItPopup)
'
'If Not (Environment.Value("blnFatalError") or reportingText = DS_SKIP) Then '*** reportingText property should be checked here
'	If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : ClickWebElementNoreportingText"
'
''	ALMvalue=Environment.Value("ALMFlag")
''	If  ALMvalue=1 Then
''		PopUpCreationtime=Cint(creationTime) + 1
''	else
'		PopUpCreationtime=Cint(creationTime)
''	End If
'
'	Dim objWebElement
'	strStepName  = "Click WebElement: "&reportingText
'
'
'	If blnIsItPopup Then
'		Set objWebElement= Browser("CreationTime:="& PopUpCreationtime).Page("Index:="& PopUpCreationtime).WebElement("class:="& className,"html tag:="& htmlTag ,"index:="& indexNumber)
'		'Set objWebElement =Browser("CreationTime:="& PopUpCreationtime).Window("text:="&windowText).Page("Index:="& PopUpCreationtime).WebElement("class:="& className,"reportingText:="&reportingText,"html tag:="& htmlTag ,"index:="& indexNumber)
'	else
'		Set objWebElement= Browser("CreationTime:="& PopUpCreationtime).Page("Index:="& PopUpCreationtime).WebElement("class:="& className,"html tag:="& htmlTag ,"index:="& indexNumber)
'	End If
'	If htmltag<> "" OR className<> "" OR indexNumber <> "" Then
'		If  objWebElement.Exist(gObjectTimeoutLimit)  then 
'			objWebElement.highlight
'			wait 2
'			objWebElement.Click
'			
'			If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
'				ImagePath=CaptureScreenshot(strStepName,objWebElement,False,False,True)
'			End if
'			strStatus = "DONE"
'			strMsg= "Web Element : "&reportingText& " clicked successfully"
'			Call ReporterFunction(strLibraryFileName,"ClickWebElementNoreportingText","2",strStepName,strMsg)
'		else
'			blnObjectError=True
'			strStatus = "FAIL"    
'			strMsg = "WebElement: " & reportingText & " Not found"
'			Call ReporterFunction(strLibraryFileName,"ClickWebElementNoreportingText","1",strStepName,strMsg) 
'			
'		end if
'	else
'		strStatus = "FAIL"
'		strMsg = "Function Parameter Not Passed Properly. Check the --ClickWebElementNoreportingText-- Function Call"
'        Call ReporterFunction(strLibraryFileName,"ClickWebElementNoreportingText","1",strStepName,strMsg )
'			
'        end if
'
'	If  blnObjectError  Then
'		Environment.Value("blnFatalError")=True
'	End If
'	If strStatus = "FAIL"  Then
'		ClickWebElementNoreportingText = strMsg
'		blnMainFailFlag = True
'		ImagePath=CaptureScreenshot(strStepName,objWebElement,False,False,True)
'	Else
'		ClickWebElementNoreportingText = True
'	End If
'	If blnDefault_eSwiftReporting Then  
'		Call UpdateResultHtml(strStepName,reportingText,strMsg,strStatus,"")
'	End If
'' //////////Word Document //////////////////////////////////////////////////////////
'	'If blnCreateTrainingDoc Then 
'	'	Check=CheckStatus_Doc(strStatus)
'	'	If Check=1 Then
'	'		Call storeCoordinates(objWebElement)
'	'		Call StoreSteps(strStepName)
'	'	End If
'	'End If
'   ' //////////Word Document ///////////
'	Set objWebElement=Nothing
'End If
'End function












'
'
'Public Function GetWebElementValue(creationTime,htmlId,className,htmltag,textPosition,Index,excelColumnName)
'If Not (Environment.Value("blnFatalError") or excelColumnName = DS_SKIP) Then
'   If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : GetWebElementValue"
'ALMvalue=getEnvironmentVariable("ALMFlag")
'If  ALMvalue=1 Then
'	PopUpCreationtime=Cint(creationTime) + 1
'else
'	PopUpCreationtime=Cint(creationTime)
'End If
'Dim objWebElement
'If htmlId<> ""  and className<>""Then
'strStepName = "Get WebElement Content "
'Set objWebElement =Browser("CreationTime:="& PopUpCreationtime).Page("Index:="& PopUpCreationtime).WebElement("html id:="& htmlId,"class:="&className,"html tag:="&htmltag,"index:="&Index)
'On error resume next 
' If  objWebElement.Exist  then 
'	 objWebElement.highlight
'     webElementMsg=objWebElement.GetROProperty("outertext")
'     MyArray=Split(webElementMsg," ","-1","1")
'	 textPosition=Cint(textPosition)
'	 val=MyArray(textPosition)
'     last  = Right(val,1)
'     If last = "-" Then
'	     finallast=instr(1,val,"-",1)
'         finalval=left(val,finallast-1)
'         val=finalval
'     End If
'	 
'               If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
'                   ImagePath=CaptureScreenshot(strStepName,objWebElement,False,False,False)
'               End if
'               
'               strStatus = "DONE"
'               strMsg= "value captured : "& val
'               Call ReporterFunction(strLibraryFileName,"GetWebElementValue","2",strStepName,strMsg)
'  else
'                blnObjectError=True
'              
'              strStatus = "FAIL"    
'              strMsg = "Web Element  object Missing"
'              Call ReporterFunction(strLibraryFileName,"GetWebElementValue","1",strStepName,strMsg)
'  end if
'else
'			  
'              strStatus = "FAIL"
'              strMsg = "Function Parameter Not Passed Properly. Check the --GetWebElementValue -- Function Call"
'              Call ReporterFunction(strLibraryFileName,"GetWebElementValue","1",strStepName,strMsg)
'     end if
'If  blnObjectError  Then
'       Environment.Value("blnFatalError")=True
'End If
'
'If strStatus = "FAIL"  Then
'       GetWebElementValue = strMsg
'       blnMainFailFlag = True
'       ImagePath=CaptureScreenshot(strStepName,objWebElement,False,False,False)
'Else
'       GetWebElementValue = val
'End If
'If blnWriteDataToOutputSheet Then
'	                                    strStepName = "Retrieve '"&val&"' value in table '"&gstrOutputSheetName&"' sheet under column "&excelColumnName
'										call WriteRunTimeDataToExcel (excelColumnName,val)
'								   ELSE
'										strStepName = "Retrieve '"&val&"' value in table"
'	 End If
'If blnDefault_eSwiftReporting Then  
'       Call UpdateResultHtml(strStepName,val,strMsg,strStatus,"")
'End If
'
'   Set objWebElement=Nothing
'   End If
'End Function
'
'Public Function VerifyWebElementById(creationTime,windowText,htmltag,innertext,HtmlId,indexNumber,blnIsItPopup)
'	If Not (Environment.Value("blnFatalError") or innertext = DS_SKIP) Then
'		If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : VerifyWebElementById"
'		
'		'ALMvalue=getEnvironmentVariable("ALMFlag")
'		'If  ALMvalue=1 Then
'			'PopUpCreationtime=Cint(creationTime) + 1
'		'else
'			PopUpCreationtime=Cint(creationTime)
'		'End If
'		Dim objWebElement
'		strStepName = "Verify if WebElement: " & innertext & " exists in the page"
'
'		If innertext<> "" OR htmltag<> "" Then
'			If blnIsItPopup Then
'				Set objWebElement = Browser("CreationTime:="& PopUpCreationtime).Window("text:="&windowText).Page("Index:="& PopUpCreationtime).WebElement("html id:="& HtmlId,"innertext:="&innertext,"html tag:="& htmlTag ,"index:="& indexNumber)
'			else
'    			Set objWebElement = Browser("CreationTime:="& PopUpCreationtime).Page("Index:="& PopUpCreationtime).WebElement("html id:="& HtmlId,"innertext:="&innertext,"html tag:="& htmlTag ,"index:="& indexNumber)
'			End If
'    		If  objWebElement.Exist(gObjectTimeoutLimit)  then 
'				objWebElement.highlight
'             	If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
'					ImagePath=CaptureScreenshot(strStepName,objWebElement,False,False,False)
'			    End if
'	            
'	            strStatus = "DONE"
'	            strMsg= "Web Element: "&innertext& " found in the page"
'	            Call ReporterFunction(strLibraryFileName,"VerifyWebElementById","2",strStepName,strMsg)
'      		else
'			blnObjectError=True
'      			
'				strStatus = "FAIL"	
'				strMsg = innertext&" Web Element Not found in the page"
'				Call ReporterFunction(strLibraryFileName,"VerifyWebElementById","1",strStepName,strMsg)
'                
'    		end if
'  		else
'     		
'     		
'            strStatus = "FAIL"
'            strMsg = "Function Parameter Not Passed Properly. Check the -VerifyWebElementById-- Function Call"
'            Call ReporterFunction(strLibraryFileName,"VerifyWebElementById","1",strStepName,strMsg)
'			
'  		End if
'  		
'  		
'	If  blnObjectError  Then
'		Environment.Value("blnFatalError")=True
'	End If
'
'	If strStatus = "FAIL"  Then
'		VerifyWebElementById = strMsg
'		blnMainFailFlag = True
'		ImagePath=CaptureScreenshot(strStepName,objWebElement,False,False,False)
'	Else
'		VerifyWebElementById = True
'	End If
'
'	If blnDefault_eSwiftReporting Then  
'		Call UpdateResultHtml(strStepName,innertext,strMsg,strStatus,"")
'	End If
'
'' //////////Word Document //////////////////////////////////////////////////////////
''	If blnCreateTrainingDoc Then
''		Check=CheckStatus_Doc(strStatus)
''		If Check=1 Then
''			Call storeCoordinates(objWebElement)
''			Call StoreSteps(strStepName)
''		End If
''	End If
'	 ' //////////Word Document ///////////
'	Set objWebElement=Nothing
'	
'	End If
'
'End function
'
'Public Function GetValueWebElementNoInnerText(creationTime,htmlId,htmltag,textPosition,Index,excelColumnName)
'If Not (Environment.Value("blnFatalError") or excelColumnName = DS_SKIP) Then
'   If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : GetValueWebElementNoInnerText"
'ALMvalue=getEnvironmentVariable("ALMFlag")
'If  ALMvalue=1 Then
'	PopUpCreationtime=Cint(creationTime) + 1
'else
'	PopUpCreationtime=Cint(creationTime)
'End If
'Dim objWebElement
'If htmlId<> "" Then
'strStepName = "Get WebElement Content "
'Set objWebElement =Browser("CreationTime:="& PopUpCreationtime).Page("Index:="& PopUpCreationtime).WebElement("html id:="& htmlId,"html tag:="&htmltag,"index:="&Index)
'On error resume next 
' If  objWebElement.Exist  then 
'	 objWebElement.highlight
'     webElementMsg=objWebElement.GetROProperty("value")
'     MyArray=Split(webElementMsg," ","-1","1")
'	 textPosition=Cint(textPosition)
'	 val=MyArray(textPosition)
'     last  = Right(val,1)
'     If last = "-" Then
'	     finallast=instr(1,val,"-",1)
'         finalval=left(val,finallast-1)
'         val=finalval
'     End If
'	 
'               If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
'                   ImagePath=CaptureScreenshot(strStepName,objWebElement,False,False,False)
'               End if
'               
'               strStatus = "DONE"
'               strMsg= "value captured : "& val
'               Call ReporterFunction(strLibraryFileName,"GetValueWebElementNoInnerText","2",strStepName,strMsg)
'  else
'                blnObjectError=True
'              
'              strStatus = "FAIL"    
'              strMsg = "Web Element  object Missing"
'              Call ReporterFunction(strLibraryFileName,"GetValueWebElementNoInnerText","1",strStepName,strMsg)
'  end if
'else
'			  
'              strStatus = "FAIL"
'              strMsg = "Function Parameter Not Passed Properly. Check the --GetValueWebElementNoInnerText -- Function Call"
'              Call ReporterFunction(strLibraryFileName,"GetValueWebElementNoInnerText","1",strStepName,strMsg)
'     end if
'If  blnObjectError  Then
'       Environment.Value("blnFatalError")=True
'End If
'
'If strStatus = "FAIL"  Then
'       GetValueWebElementNoInnerText = strMsg
'       blnMainFailFlag = True
'       ImagePath=CaptureScreenshot(strStepName,objWebElement,False,False,False)
'Else
'       GetValueWebElementNoInnerText = True
'End If
'If blnWriteDataToOutputSheet Then
'	                                    strStepName = "Retrieve '"&val&"' value in table '"&gstrOutputSheetName&"' sheet under column "&excelColumnName
'										call WriteRunTimeDataToExcel (excelColumnName,val)
'								   ELSE
'										strStepName = "Retrieve '"&val&"' value in table"
'	 End If
'If blnDefault_eSwiftReporting Then  
'       Call UpdateResultHtml(strStepName,val,strMsg,strStatus,"")
'End If
'
'   Set objWebElement=Nothing
'   End If
'End Function
'
'Public Function SetWebEditAndSubmit(creationTime,windowText,webeditName,webeditTexttype,webeditIndex,val,blnIsItPopup)
'	If Not (Environment.Value("blnFatalError") or val = DS_SKIP) Then
'		If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : SetWebEditAndSubmit"
'
'		ALMvalue=getEnvironmentVariable("ALMFlag")
'		If  ALMvalue=1 Then
'			PopUpCreationtime=Cint(creationTime) + 1
'		else
'			PopUpCreationtime=Cint(creationTime)
'		End If
'
'		Dim objWebEdit
'		strStepName = "Set text box : "& webeditName
'		If webeditName<> ""  or webeditTexttype <> "" or webeditIndex <> "" Then
'
'			If blnIsItPopup Then
'				Set objWebEdit = Browser("CreationTime:="& PopUpCreationtime).Window("text:="&windowText).Page("Index:="& PopUpCreationtime).WebEdit("name:="&webeditName,"type:=" &webeditTexttype,"index:="&webeditIndex)
'			else
'    				Set objWebEdit = Browser("CreationTime:="& PopUpCreationtime).Page("Index:="& PopUpCreationtime).WebEdit("name:="&webeditName,"type:=" &webeditTexttype,"index:="&webeditIndex)
'			End If
'			If  objWebEdit.Exist(gObjectTimeoutLimit)  then 
'     				objWebEdit.Click
'   	 			objWebEdit.Set val
'				objWebEdit.Submit
'	 			Wait(1)
'      			If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
'					ImagePath=CaptureScreenshot(strStepName,objWebEdit,False,False,True)
'	  			End if
'				
'	            strStatus = "DONE"
'	            strMsg= "WebEdit value set successfully"
'	            Call ReporterFunction(strLibraryFileName,"SetWebEditAndSubmit","2",strStepName,strMsg)
'  			else
'  				blnObjectError=True
'               	strStatus = "FAIL"
'               	strMsg = "Object Missing--Check the Screen Properly"
'               	Call ReporterFunction(strLibraryFileName,"SetWebEditAndSubmit","1",strStepName,strMsg)
'  			end if
'		else
'               	strStatus = "FAIL"
'               	strMsg = "Function Parameter Not Passed Properly. Check the -SetWebEditAndSubmit-- Function Call"
'               	Call ReporterFunction(strLibraryFileName,"SetWebEditAndSubmit","1",strStepName,strMsg)
'		end if
'
'	If  blnObjectError  Then
'		Environment.Value("blnFatalError")=True
'	End If
'
'	If strStatus = "FAIL"  Then
'		SetWebEditAndSubmit = strMsg
'		blnMainFailFlag = True
'		ImagePath=CaptureScreenshot(strStepName,objWebEdit,False,False,True)
'	Else
'		SetWebEditAndSubmit = True
'	End If
'
'	If blnDefault_eSwiftReporting Then  
'		Call UpdateResultHtml(strStepName,val,strMsg,strStatus,"")
'	End If
'
'	Set objWebEdit=Nothing
'	
'	End If
'End function
'Public Function SendWebKey(keyName)
'
'If Not (Environment.Value("blnFatalError") or name = DS_SKIP) Then '*** name property should be checked here
'	If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : SendWebKey"
'
'
'	strStepName  = "SendWebKey: "&keyName
'
'
'If keyName <> "" Then
'Set ObjSendKey= CreateObject("WScript.Shell")
'ObjSendKey.SendKeys "{"&keyName&"}"
'strStatus = "DONE"
'strMsg= ""&keyName& " pressed successfully"
'Call ReporterFunction(strLibraryFileName,"SendWebKey","2",strStepName,strMsg)
'else
'		strStatus = "FAIL"
'		strMsg= ""&keyName& " is empty"
'Call ReporterFunction(strLibraryFileName,"SendWebKey","2",strStepName,strMsg)
'end if
'If  blnObjectError  Then
'		Environment.Value("blnFatalError")=True
'	End If
'
'If strStatus = "FAIL"  Then
'		SendWebKey = strMsg
'		blnMainFailFlag = True
'	
'	Else
'		SendWebKey = True
'	End If
'	If blnDefault_eSwiftReporting Then  
'		Call UpdateResultHtml(strStepName,keyName,strMsg,strStatus,"")
'	End If
'
'	Set ObjSendKey= Nothing
'
'End If
'
'End function
'
' Public Function VerifyWebLinkOutertext(creationTime,windowText,htmltag,outertext,indexNumber,blnIsItPopup)
'	If Not (Environment.Value("blnFatalError") or outertext = DS_SKIP) Then
'		If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : VerifyWebLinkOutertext"
'		
'		ALMvalue=getEnvironmentVariable("ALMFlag")
'		If  ALMvalue=1 Then
'			PopUpCreationtime=Cint(creationTime) + 1
'		else
'			PopUpCreationtime=Cint(creationTime)
'		End If
'		Dim objWebLink
'		strStepName = "Verify if Link: " & outertext & " exists in the page"
'
'		If outertext<> "" OR htmltag<> "" Then
'			If blnIsItPopup Then
'				Set objWebLink = Browser("CreationTime:="& PopUpCreationtime).Window("text:="&windowText).Page("Index:="& PopUpCreationtime).Link("outertext:="&outertext,"html tag:="& htmlTag ,"index:="& indexNumber)
'			else
'    			Set objWebLink = Browser("CreationTime:="& PopUpCreationtime).Page("Index:="& PopUpCreationtime).Link("outertext:="&outertext,"html tag:="& htmlTag ,"index:="& indexNumber)
'			End If
'    		If  objWebLink.Exist(gObjectTimeoutLimit)  then 
'				objWebLink.highlight
'             	If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
'					ImagePath=CaptureScreenshot(strStepName,objWebElement,False,False,False)
'			    End if
'	            
'	            strStatus = "DONE"
'	            strMsg= "Web Link: "&outertext& " found in the page"
'	            Call ReporterFunction(strLibraryFileName,"VerifyWebLinkOutertext","2",strStepName,strMsg)
'      		else
'			blnObjectError=True
'      			
'				strStatus = "FAIL"	
'				strMsg = outertext&" Web Link Not found in the page"
'				Call ReporterFunction(strLibraryFileName,"VerifyWebLinkOutertext","1",strStepName,strMsg)
'                
'    		end if
'  		else
'     		
'     		
'            strStatus = "FAIL"
'            strMsg = "Function Parameter Not Passed Properly. Check the -VerifyWebLinkOutertext-- Function Call"
'            Call ReporterFunction(strLibraryFileName,"VerifyWebLinkOutertext","1",strStepName,strMsg)
'			
'  		End if
'  		
'  		
'	If  blnObjectError  Then
'		Environment.Value("blnFatalError")=True
'	End If
'
'	If strStatus = "FAIL"  Then
'		VerifyWebLinkOutertext = strMsg
'		blnMainFailFlag = True
'		ImagePath=CaptureScreenshot(strStepName,objWebElement,False,False,False)
'	Else
'		VerifyWebLinkOutertext = True
'	End If
'
'	If blnDefault_eSwiftReporting Then  
'		Call UpdateResultHtml(strStepName,outertext,strMsg,strStatus,"")
'	End If
'
'
'	Set objWebElement=Nothing
'	
'	End If
'
'End function
'''//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
''Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
''Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet Call
'
''**************************************************************************************************************************
'

