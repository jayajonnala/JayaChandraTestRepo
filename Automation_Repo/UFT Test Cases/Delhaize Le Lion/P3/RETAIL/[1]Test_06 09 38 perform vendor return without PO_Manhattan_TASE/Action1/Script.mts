
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


gstrTestCaseName = "Test_06 09 38 perform ven without PO"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\TASEWork\Data\TASE_DT_114_SRM Standard SC Free Text on Cost Cent P1_.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''''''Call WriteRunTimeDataToExcelGlobalSheet ("DT_TIME",Cstr(Hour(now())&":"&Minute(Now())))
''''''''''''--------------login----------------
''''''''''''''''Close All Browser
Call CloseAllBrowsers()
''''Launch Manhattan URL
Call LaunchSAPWebApplicationEdge(DT_URL)
wait 30
'Login 
Call LoginManhattan(0,DT_USER,DT_PASSWORD)
wait 10
''''Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''click hamburger icon
Call ClickWebElementNoreportingText(0,"","SPAN","Main Menu","x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-topbar-menu-icon",0,False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'Click Show All button
Call ClickWebElement(0,"","SPAN","Show All","x-btn-inner x-btn-inner-default-medium",0,False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''Click Ditribution tab
Call ClickWebElement(0,"","SPAN","Distribution","x-tab-inner x-tab-inner-default",1,False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''Click Distribution Order
Call ClickWebElement(0,"","SPAN","Distribution Orders","",0,False)
wait 5
''PRIMARY FIELD
'wait 5
Call SetWebEdit(0, "", "combobox.*inputEl", "text", 2, DT_DO_COMBO_KEY, False)
'''Enter the DO number
Call SetWebEdit(0, "", "DistributionorderID", "text", 0, DT_DO, False)
'Click Apply
Call ClickWebElement(0,"","SPAN","Apply","x-btn-inner x-btn-inner-default-small",0,False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'Verify DO status
Call VerifyWebElement(0,"","DIV",DT_DO_STATUS,"x-grid-cell-inner",0,False)
'Select CheckBox
Call ClickWebElement(0, "", "DIV", "", "x-grid-row-checker", 0, False)
'Click Edit dates
Call ClickWebElement(0, "", "SPAN", "Edit Dates", "x-btn-inner x-btn-inner-default-small", 0, False)
''''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
''Enter pickup start
Call SetWebEdit(0, "", "pickupStartDTTMString03", "text", 0, Cstr(DT_PICKUP_START), False)
''Enter pickup end
Call SetWebEdit(0, "", "pickupEndDTTMString04", "text", 0,Cstr( DT_PICKUP_END), False)
''Enter delivery start
Call SetWebEdit(0, "", "deliveryStartDTTMString05", "text", 0, Cstr(DT_DELIVERY_START), False)
''Enter delivery end
Call SetWebEdit(0, "", "deliveryEndDTTMString06", "text", 0, Cstr(DT_DELIVERY_END), False)
''click save
Call ClickWebButton(0, "", "btn", "Save", "BUTTON", 0, False)
'''click hamburger icon
Call ClickWebElementNoreportingText(0,"","SPAN","Main Menu","x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-topbar-menu-icon",0,False)
'Click Show All button
Call ClickWebElement(0,"","SPAN","Show All","x-btn-inner x-btn-inner-default-medium",0,False)
'Click Configuration tab
Call ClickWebElement(0, "", "SPAN", "Configuration", "x-tab-inner x-tab-inner-default", 1, False)
wait 5
'Click Items
Call ClickWebElement(0,"","SPAN","Items","",1,False)
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
''GET Allocation type
Call GetWebElementValue(0,"dataForm:ItemFacilityDetails_SOMAllocationType","captionData","SPAN",0,0,"DT_ALLOCATION_TYPE_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
'''click hamburger icon
Call ClickWebElementNoreportingText(0,"","SPAN","Main Menu","x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-topbar-menu-icon",0,False)
'Click Show All button
Call ClickWebElement(0,"","SPAN","Show All","x-btn-inner x-btn-inner-default-medium",0,False)
''Click Ditribution tab
Call ClickWebElement(0,"","SPAN","Distribution","x-tab-inner x-tab-inner-default",1,False)
'Click Dock Door
Call ClickWebElement(0,"","SPAN","Dock Door","",0,False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
wait 5
''Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''''''Verify Allocation
Call VerifyWebElementById(0,"","SPAN", DT_DOCK_DOOR,"dataForm:listView:dataTable:"&DT_HTML_ID&":custId3",0,False)
''Verify Status
Call VerifyWebElementById(0,"","SPAN", DT_STATUS,"dataForm:listView:dataTable:"&DT_HTML_ID&":out5",0,False)
'''Verify Dock Door
Call VerifyWebElementById(0,"","SPAN", DT_DOCK_DOOR_VALUE,"dataForm:listView:dataTable:"&DT_HTML_ID&":out4",0,False)
''''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
''click hamburger icon
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
''Click Distribution Order
Call ClickWebElement(0,"","SPAN","Distribution Orders","",0,False)
wait 5
''PRIMARY FIELD
'wait 5
Call SetWebEdit(0, "", "combobox.*inputEl", "text", 2, DT_DO_COMBO_KEY, False)
'''Enter the DO number
Call SetWebEdit(0, "", "DistributionorderID", "text", 0, DT_DO, False)
'Click Apply
Call ClickWebElement(0,"","SPAN","Apply","x-btn-inner x-btn-inner-default-small",0,False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''Select CheckBox
Call ClickWebElement(0, "", "DIV", "", "x-grid-row-checker", 0, False)
'Click Edit Header
Call ClickWebElement(0, "", "SPAN", "Edit Header", "x-btn-inner x-btn-inner-default-small", 0, False)
''''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
'''click pickup/delivery
Call ClickLink(0, "", "tab_link", "Pickup\/Delivery", False)

'''Browser("Manhattan Associates").Page("Manhattan Associates").Frame("uxiframe-3996-frame").Link("Pickup/Delivery").Click
''enter dock door number
Call SetWebEdit(0, "", "dataForm:DOEdit_PickupDeliveryTab_InText_OriginDockdoor", "text", 0, DT_DOCK_DOOR_VALUE, False)
''click save
Call ClickWebButton(0, "", "btn", "Save", "INPUT", 0, False)
''''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
'''click hamburger icon
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
''Click Distribution Order
Call ClickWebElement(0,"","SPAN","Distribution Orders","",0,False)
wait 5
''PRIMARY FIELD
'wait 5
Call SetWebEdit(0, "", "combobox.*inputEl", "text", 2, DT_DO_COMBO_KEY, False)
'''Enter the DO number
Call SetWebEdit(0, "", "DistributionorderID", "text", 0, DT_DO, False)
'Click Apply
Call ClickWebElement(0,"","SPAN","Apply","x-btn-inner x-btn-inner-default-small",0,False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''Select CheckBox
Call ClickWebElement(0, "", "DIV", "", "x-grid-row-checker", 0, False)
''Click More
Call ClickWebElement(0, "", "SPAN", "More", "x-btn-inner x-btn-inner-default-small", 0, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''click wave
Call ClickWebElement(0, "", "SPAN", "Wave", "x-menu-item-text x-menu-item-text-default x-menu-item-indent", 0, False)
'''''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
''''select wave type
Call SelectWebCheckbox(0, "", "checkAll_c"&DT_HTML_ID_WAVE_TEMPLATE&"_dataForm:listView:dataTable", "INPUT", 0, "ON", False)
' ''click Run wave
 Call ClickWebButton(0, "", "btn", "Run Wave", "INPUT", 0, False)
''Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'' ''click submit wave
Call ClickWebButton(0, "", "btn", "Submit", "INPUT", 0, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'Click on the wave number
waveNum=Cstr(DT_WAVE_NUMBER+".*")
Call ClickLink(0, "", "", waveNum, blnIsItPopup)
wait 5
''Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'''''''verify wave status
'''''''''''Call VerifyWebElement(0, "", "SPAN", DT_WAVE_STATUS, "", 0, False)
While IfWebElementExists(0, "", "SPAN", DT_WAVE_STATUS, "", 0, False)<> True
	''Call ClickWebElement(0, "", "INPUT", "", "paginationCtrlCls", 0, False)
	Call ClickWebElementNoreportingText(0,"","INPUT","Refresh","paginationCtrlCls", 0, False)
	wait 10
Wend
'''select wave checkbox
Call SelectWebCheckbox(0, "", "checkAll_c0_dataForm:listView:dataTable", "INPUT", 0, "ON", False)
'' click view
Call ClickWebButton(0, "", "btn", "View", "INPUT", 0, False)
'''Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''lines summary - required
Call GetValueWebElementInnerText(0,"dataForm:LinesRequired","SPAN",0,0,"DT_LINES_REQUIRED_OUTPUT")
''lines summary - allocated
Call GetValueWebElementInnerText(0,"dataForm:LinesAllocated","SPAN",0,0,"DT_LINES_ALLOCATED_OUTPUT")
''units summary - required
Call GetValueWebElementInnerText(0,"dataForm:UnitsRequired","SPAN",0,0,"DT_UNITS_REQUIRED_OUTPUT")
''units summary - allocated
Call GetValueWebElementInnerText(0,"dataForm:UnitsAllocated","SPAN",0,0,"DT_UNITS_ALLOCATED_OUTPUT")
''''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
''''''click hamburger icon
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
''Click Distribution Order
Call ClickWebElement(0,"","SPAN","Distribution Orders","",0,False)
wait 5
''PRIMARY FIELD
'wait 5
Call SetWebEdit(0, "", "combobox.*inputEl", "text", 2, DT_DO_COMBO_KEY, False)
'''Enter the DO number
Call SetWebEdit(0, "", "DistributionorderID", "text", 0, DT_DO, False)
'Click Apply
Call ClickWebElement(0,"","SPAN","Apply","x-btn-inner x-btn-inner-default-small",0,False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'Verify DO status
'''Call VerifyWebElement(0,"","DIV",DT_DO_STATUS_1,"x-grid-cell-inner",0,False)
While  IfWebElementExists(0,"","DIV",DT_DO_STATUS_1,"x-grid-cell-inner",0,False)<> True
	Call ClickWebElementNoreportingText(0, "", "SPAN", "Refresh", "x-btn-icon-el x-btn-icon-el-default-toolbar-small x-tbar-loading", 0, False)
	wait 10
Wend
'Select CheckBox
Call ClickWebElement(0, "", "DIV", "", "x-grid-row-checker", 0, False)
''Click View
Call ClickWebElement(0, "", "SPAN", "View", "x-btn-inner x-btn-inner-default-small", 0, False)
''Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'''''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
'''GET parent ordr id
Call GetValueWebElementInnerText(0,"dataForm:DODetailsHeader_OutText_ParentorderID","SPAN",0,0,"DT_PARENT_ID_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'click parent order id
Call ClickWebElement(0, "", "SPAN", DT_PARENT_ID, "", 0, False)
''Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''get shipment id
Call GetValueWebElementInnerText(0,"dataForm:shipmentIdRepeat1:0:DODtlHdr_ShpId_OpLnk_Txt__","SPAN",0,0,"DT_SHIPMENT_ID_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''''''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
''''''click hamburger icon
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
''Click Shipments
Call ClickWebElement(0,"","SPAN","Shipments","",0,False)
wait 5
'''PRIMARY FIELD
''wait 5
Call SetWebEdit(0, "", "combobox.*inputEl", "text", 2, DT_SHIPMENT_COMBO_KEY, False)
'''''Enter the Shipment ID
Call SetWebEdit(0, "", "TCShipmentID", "text", 0, DT_SHIPMENT_ID, False)
''Click Apply
Call ClickWebElement(0,"","SPAN","Apply","x-btn-inner x-btn-inner-default-small",0,False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''''Click More
Call ClickWebElement(0, "", "SPAN", "More", "x-btn-inner x-btn-inner-default-small", 0, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''Select CheckBox
Call ClickWebElement(0, "", "DIV", "", "x-grid-row-checker", 0, False)
''click Close
Call ClickWebElement(0, "", "SPAN", "Close", "x-menu-item-text x-menu-item-text-default x-menu-item-indent", 0, False)
'''''Close the screen
Call ClickWebElementNoreportingText(0, "", "SPAN", "Open Windows", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-window", 0, False)
Call DoubleClickWebElement(0, "", "SPAN", "Close", "x-btn-inner x-btn-inner-default-small", 0, False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''set lpn variance
Call SelectWebList(0, "", "dataForm:closeShipmentTable:0:lpnProcessOption", 0, DT_LPN_VARIANCE, False)
''click save
Call ClickWebButton(0, "", "btn", "Save", "INPUT", 0, False)
'''''click hamburger icon
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
''Click Distribution Order
Call ClickWebElement(0,"","SPAN","Distribution Orders","",0,False)
wait 5
''PRIMARY FIELD
'wait 5
Call SetWebEdit(0, "", "combobox.*inputEl", "text", 2, DT_DO_COMBO_KEY, False)
'''Enter the DO number
Call SetWebEdit(0, "", "DistributionorderID", "text", 0, DT_DO, False)
'Click Apply
Call ClickWebElement(0,"","SPAN","Apply","x-btn-inner x-btn-inner-default-small",0,False)
'Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
'Verify DO status
Call VerifyWebElement(0,"","DIV",DT_DO_STATUS_2,"x-grid-cell-inner",0,False)
''''Capture Screenshot
Call CaptureWebScreenNew(0,"Capture Screen")
''sign out
Call ClickWebElement(0, "", "SPAN", "", "x-btn-icon-el x-btn-icon-el-default-toolbar-medium wt-user ", 0, False)
Call ClickWebElement(0, "", "SPAN","Sign out", "x-btn-inner x-btn-inner-default-small", 0, False)
Call CloseBrowser(0)
Call FinalStatus()













''Login for Manhattan
'Public Function LoginManhattan(creationTime,UserName,Password)
'If Not (Environment.Value("blnFatalError") or buttonName = DS_SKIP) Then
'	If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : Login Manhattan Active"
'		PopUpCreationtime=Cint(creationTime)
'	Dim objButton
'	strStepName = "Enter User Name : "&UserName&" and Password : ********"
'	strInputValue= UserName&" ********"
'	Set obj=Browser("creationtime:="&creationTime).Page("creationtime:="&creationTime)
'	 If  obj.Exist  then 
'	 	If UserName<>"" and Password <> "" Then
'	 		Set objUser=obj.WebEdit("html id:=username")
'			objUser.Set UserName
'		  	obj.WebEdit("html id:=password").Set Password	
'			Wait 5
'			obj.WebButton("html id:=loginButton").ClicK
'			wait(5)
'			If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
'				ImagePath=CaptureScreenshot(strStepName,objUser,True,True,True)
'			End if
'			strStatus = "DONE"
'			strMsg= "Login to Manhattan Active application successfull"
'			Call ReporterFunction(strLibraryFileName,"Login Manhattan Active","2",strStepName,strMsg)
'		else
'			blnObjectError=True
'			strStatus = "FAIL"    
'			strMsg = "Manhattan Active application Login Page Not Found"
'			Call ReporterFunction(strLibraryFileName,"Login Manhattan Active","1",strStepName,strMsg)
'		end if
'	else
'		
'		strStatus = "FAIL"
'		strMsg = "Function Parameter Not Passed Properly. Check the -Login Manhattan Active-- Function Call"
'		Call ReporterFunction(strLibraryFileName,"Login Manhattan Active","1",strStepName,strMsg)
'	end if
'	If  blnObjectError  Then
'		Environment.Value("blnFatalError")=True
'	End If
'	
'	If blnDefault_eSwiftReporting Then  
'		Call UpdateResultHtml(strStepName,strInputValue,strMsg,strStatus,"")
'	End If
'	' //////////Word Document //////////////////////////////////////////////////////////
'	'If blnCreateTrainingDoc Then 
'	'	Check=CheckStatus_Doc(strStatus)
'	'	If Check=1 Then
'	'		Call storeCoordinates(objButton)
'	'		Call StoreSteps(strStepName)
'	'	End If
'	'End If
'	   ' //////////Word Document ///////////
'	Set obj=Nothing
'End If
'End function
'
'Public Function ClickWebElementById(creationTime,windowText,htmltag,HtmlId,className,indexNumber,blnIsItPopup)
'
'If Not (Environment.Value("blnFatalError") or HtmlId = DS_SKIP) Then '*** HtmlId property should be checked here
'	If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : ClickWebElementById"
'
''	ALMvalue=Environment.Value("ALMFlag")
''	If  ALMvalue=1 Then
''		PopUpCreationtime=Cint(creationTime) + 1
''	else
'		PopUpCreationtime=Cint(creationTime)
''	End If
'
'	Dim objWebElement
'	strStepName  = "Click WebElement: "&HtmlId
'
'
'	If blnIsItPopup Then
'		Set objWebElement= Browser("CreationTime:="& PopUpCreationtime).Page("Index:="& PopUpCreationtime).WebElement("class:="& className,"html id="&HtmlId,"html tag:="& htmlTag ,"index:="& indexNumber)
'		'Set objWebElement =Browser("CreationTime:="& PopUpCreationtime).Window("text:="&windowText).Page("Index:="& PopUpCreationtime).WebElement("class:="& className,"HtmlId:="&HtmlId,"html tag:="& htmlTag ,"index:="& indexNumber)
'	else
'		Set objWebElement= Browser("CreationTime:="& PopUpCreationtime).Page("Index:="& PopUpCreationtime).WebElement("class:="& className,"html id:="&HtmlId,"html tag:="& htmlTag ,"index:="& indexNumber)
'	End If
'	If HtmlId<> "" OR htmltag<> "" OR className<> "" OR indexNumber <> "" Then
'		If  objWebElement.Exist(gObjectTimeoutLimit)  then 
'			objWebElement.highlight
'			wait 2
'			objWebElement.Click
'			If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
'				ImagePath=CaptureScreenshot(strStepName,objWebElement,False,False,True)
'			End if
'			strStatus = "DONE"
'			strMsg= "Web Element : "&HtmlId& " clicked successfully"
'			Call ReporterFunction(strLibraryFileName,"ClickWebElementById","2",strStepName,strMsg)
'		else
'			blnObjectError=True
'			strStatus = "FAIL"    
'			strMsg = "WebElement: " & HtmlId & " Not found"
'			Call ReporterFunction(strLibraryFileName,"ClickWebElementById","1",strStepName,strMsg) 
'			
'		end if
'	else
'		strStatus = "FAIL"
'		strMsg = "Function Parameter Not Passed Properly. Check the --ClickWebElementById-- Function Call"
'        Call ReporterFunction(strLibraryFileName,"ClickWebElementById","1",strStepName,strMsg )
'			
'        end if
'
'	If  blnObjectError  Then
'		Environment.Value("blnFatalError")=True
'	End If
'	If strStatus = "FAIL"  Then
'		ClickWebElementById = strMsg
'		blnMainFailFlag = True
'		ImagePath=CaptureScreenshot(strStepName,objWebElement,False,False,True)
'	Else
'		ClickWebElementById = True
'	End If
'	If blnDefault_eSwiftReporting Then  
'		Call UpdateResultHtml(strStepName,HtmlId,strMsg,strStatus,"")
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
'     If webElementMsg="" Then
'     	webElementMsg=objWebElement.GetROProperty("outertext")
'     End If
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
'     			objWebEdit.Click
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
''//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet Call

'**************************************************************************************************************************

