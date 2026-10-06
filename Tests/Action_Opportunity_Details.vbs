' --- Config ------------------------------------------------------------------
Const APP_URL               = "https://your-org--qa.sandbox.lightning.force.com/"  ' TODO: set to your actual sandbox/org URL
Const APP_BROWSER           = "msedge"
Const APP_BROWSER_ARGS      = "-inprivate"
Const DEFAULT_SYNC_MS       = 30000
Const DEFAULT_PAGE_LOAD_MS  = 60000
Const BROWSER_TITLE_PATTERN = ".*Salesforce.*|.*Lightning.*"
Const PAGE_TITLE_PATTERN    = ".*Salesforce.*|.*Lightning.*"

' --- Utilities -----------------------------------------------------------------

Function GetAppBrowser()
    Dim oDesc
    Set oDesc = Description.Create()
    oDesc("title").Value = BROWSER_TITLE_PATTERN
    oDesc("title").RegularExpression = True
    Set GetAppBrowser = Browser(oDesc)
End Function

Function GetAppPage()
    Dim oDesc
    Set oDesc = Description.Create()
    oDesc("title").Value = PAGE_TITLE_PATTERN
    oDesc("title").RegularExpression = True
    Set GetAppPage = GetAppBrowser().Page(oDesc)
End Function

Function WaitForObject(oTestObject, nTimeoutMs)
    Dim nTimeoutSec
    If nTimeoutMs = 0 Then nTimeoutMs = DEFAULT_SYNC_MS
    nTimeoutSec = nTimeoutMs / 1000
    WaitForObject = oTestObject.Exist(nTimeoutSec)
End Function

Sub ReportStep(sStepName, bPassed, sDetails)
    Dim nStatus
    If bPassed Then
        nStatus = micPass
    Else
        nStatus = micFail
    End If
    Reporter.ReportEvent nStatus, sStepName, sDetails
End Sub

Sub ReportFatal(sStepName, sDetails)
    Reporter.ReportEvent micFail, sStepName, sDetails
    ExitAction
End Sub

' --- Page Object: Login (incl. two-factor Verification Code step) -------------
' Shared with Action_Opportunity.vbs - steps 1-7 are identical for both test cases.

Class PO_Login

    Private sLocUsername
    Private sLocPassword
    Private sLocBtnLogin
    Private sLocVerificationCode
    Private sLocBtnVerify

    Private Sub Class_Initialize()
        sLocUsername         = "xpath:=//input[@id='username']"
        sLocPassword         = "xpath:=//input[@id='password']"
        sLocBtnLogin         = "xpath:=//input[@id='Login']"
        sLocVerificationCode = "xpath:=//input[@id='tc']"
        sLocBtnVerify        = "xpath:=//input[@id='save']"
    End Sub

    Public Function EnterUsername(ByVal sUsername)
        Dim oPage
        Set oPage = GetAppPage()
        If WaitForObject(oPage.WebEdit(sLocUsername), DEFAULT_SYNC_MS) Then
            oPage.WebEdit(sLocUsername).Set sUsername
            EnterUsername = True
        Else
            ReportStep "PO_Login.EnterUsername", False, "Username field not found within timeout"
            EnterUsername = False
        End If
    End Function

    Public Function EnterPassword(ByVal sPassword)
        Dim oPage
        Set oPage = GetAppPage()
        If WaitForObject(oPage.WebEdit(sLocPassword), DEFAULT_SYNC_MS) Then
            oPage.WebEdit(sLocPassword).SetSecure sPassword
            EnterPassword = True
        Else
            ReportStep "PO_Login.EnterPassword", False, "Password field not found within timeout (step 1 may have failed)"
            EnterPassword = False
        End If
    End Function

    Public Function ClickLogin()
        Dim oPage
        Set oPage = GetAppPage()
        If WaitForObject(oPage.WebButton(sLocBtnLogin), DEFAULT_SYNC_MS) Then
            oPage.WebButton(sLocBtnLogin).Click
            ClickLogin = True
        Else
            ReportStep "PO_Login.ClickLogin", False, "Login button not found within timeout"
            ClickLogin = False
        End If
    End Function

    Public Function EnterVerificationCode(ByVal sCode)
        Dim oPage
        Set oPage = GetAppPage()
        If WaitForObject(oPage.WebEdit(sLocVerificationCode), DEFAULT_SYNC_MS) Then
            oPage.WebEdit(sLocVerificationCode).Set sCode
            EnterVerificationCode = True
        Else
            ReportStep "PO_Login.EnterVerificationCode", False, "Verification Code field not found within timeout"
            EnterVerificationCode = False
        End If
    End Function

    Public Function ClickVerify()
        Dim oPage
        Set oPage = GetAppPage()
        If WaitForObject(oPage.WebButton(sLocBtnVerify), DEFAULT_SYNC_MS) Then
            oPage.WebButton(sLocBtnVerify).Click
            ClickVerify = True
        Else
            ReportStep "PO_Login.ClickVerify", False, "Verify button not found within timeout"
            ClickVerify = False
        End If
    End Function

    ' Drives the full three-step flow: username -> Login (reveals password) ->
    ' password -> Login (reveals verification code) -> code -> Verify.
    Public Function Login(ByVal sUsername, ByVal sPassword, ByVal sVerificationCode)
        Dim bStep1, bStep2

        bStep1 = EnterUsername(sUsername) And ClickLogin()
        If Not bStep1 Then
            ReportStep "PO_Login.Login", False, "Step 1 (username submit) failed"
            Login = False
            Exit Function
        End If

        bStep2 = EnterPassword(sPassword) And ClickLogin()
        If Not bStep2 Then
            ReportStep "PO_Login.Login", False, "Step 2 (password submit) failed"
            Login = False
            Exit Function
        End If

        Login = EnterVerificationCode(sVerificationCode) And ClickVerify()
    End Function

    Public Function IsLoginPageDisplayed()
        Dim oPage
        Set oPage = GetAppPage()
        IsLoginPageDisplayed = WaitForObject(oPage.WebEdit(sLocUsername), DEFAULT_SYNC_MS)
    End Function

End Class

' --- Page Object: App Navigator (App Launcher + top Navigation Menu) ----------
' Shared with Action_Opportunity.vbs - steps 8-14 are identical for both test cases.

Class PO_AppNavigator

    Private sLocAppLauncherBtn
    Private sLocSearchAppsBox
    Private sLocShowNavMenuBtn

    Private Sub Class_Initialize()
        sLocAppLauncherBtn = "xpath:=//button[@title='App Launcher']"
        sLocSearchAppsBox  = "xpath:=//input[@placeholder='Search apps and items...']"
        sLocShowNavMenuBtn = "xpath:=//button[@title='Show Navigation Menu']"
    End Sub

    Public Function OpenAppLauncher()
        Dim oPage
        Set oPage = GetAppPage()
        If WaitForObject(oPage.WebButton(sLocAppLauncherBtn), DEFAULT_SYNC_MS) Then
            oPage.WebButton(sLocAppLauncherBtn).Click
            OpenAppLauncher = True
        Else
            ReportStep "PO_AppNavigator.OpenAppLauncher", False, "App Launcher button not found within timeout"
            OpenAppLauncher = False
        End If
    End Function

    ' Types the app name into the App Launcher search box and clicks the
    ' matching app tile from the filtered results.
    Public Function OpenApp(ByVal sAppName)
        Dim oPage, sLocAppOption
        Set oPage = GetAppPage()

        If Not WaitForObject(oPage.WebEdit(sLocSearchAppsBox), DEFAULT_SYNC_MS) Then
            ReportStep "PO_AppNavigator.OpenApp", False, "Search apps and items box not found within timeout"
            OpenApp = False
            Exit Function
        End If
        oPage.WebEdit(sLocSearchAppsBox).Set sAppName

        sLocAppOption = "xpath:=//*[@role='option'][contains(.,'" & sAppName & "')]"
        If Not WaitForObject(oPage.WebElement(sLocAppOption), DEFAULT_SYNC_MS) Then
            ReportStep "PO_AppNavigator.OpenApp", False, "App option '" & sAppName & "' not found within timeout"
            OpenApp = False
            Exit Function
        End If
        oPage.WebElement(sLocAppOption).Click
        OpenApp = True
    End Function

    Public Function ShowNavigationMenu()
        Dim oPage
        Set oPage = GetAppPage()
        If WaitForObject(oPage.WebButton(sLocShowNavMenuBtn), DEFAULT_SYNC_MS) Then
            oPage.WebButton(sLocShowNavMenuBtn).Click
            ShowNavigationMenu = True
        Else
            ReportStep "PO_AppNavigator.ShowNavigationMenu", False, "Show Navigation Menu button not found within timeout"
            ShowNavigationMenu = False
        End If
    End Function

    ' Clicks a menu item (e.g. "Opportunities") from the open Navigation Menu.
    Public Function ClickNavMenuItem(ByVal sItemName)
        Dim oPage, sLocMenuItem
        Set oPage = GetAppPage()
        sLocMenuItem = "xpath:=//*[@role='menuitem'][contains(.,'" & sItemName & "')]"
        If WaitForObject(oPage.WebElement(sLocMenuItem), DEFAULT_SYNC_MS) Then
            oPage.WebElement(sLocMenuItem).Click
            ClickNavMenuItem = True
        Else
            ReportStep "PO_AppNavigator.ClickNavMenuItem", False, "Menu item '" & sItemName & "' not found within timeout"
            ClickNavMenuItem = False
        End If
    End Function

    ' Verifies the app shell heading (top-left, e.g. "Sales Centre") is displayed.
    Public Function IsAppDisplayed(ByVal sAppName)
        Dim oPage, sLocHeading
        Set oPage = GetAppPage()
        sLocHeading = "xpath:=//h1[text()='" & sAppName & "']"
        IsAppDisplayed = WaitForObject(oPage.WebElement(sLocHeading), DEFAULT_PAGE_LOAD_MS)
    End Function

End Class

' --- Page Object: Opportunity List -----------------------------------------
' Shared with Action_Opportunity.vbs - steps 15-20 are identical for both test cases.

Class PO_OpportunityList

    Private sLocSelectListViewBtn
    Private sLocSearchListsBox
    Private sLocSearchThisListBox

    Private Sub Class_Initialize()
        ' NOTE: the inner <button title="Select a List View: ..."> only
        ' highlights/focuses on click - the dropdown's open handler lives on
        ' the enclosing <lightning-button-icon class="pickerChevron"> custom
        ' element, so that is the actual click target (confirmed 2026-10-06).
        sLocSelectListViewBtn = "xpath:=//lightning-button-icon[@class='pickerChevron']"
        sLocSearchListsBox    = "xpath:=//input[@placeholder='Search lists...']"
        sLocSearchThisListBox = "xpath:=//input[@placeholder='Search this list...']"
    End Sub

    Public Function IsOpportunitiesPageDisplayed()
        Dim oPage, sLocHeading
        Set oPage = GetAppPage()
        sLocHeading = "xpath:=//h1[text()='Opportunities']"
        IsOpportunitiesPageDisplayed = WaitForObject(oPage.WebElement(sLocHeading), DEFAULT_PAGE_LOAD_MS)
    End Function

    ' Opens the list-view picker, types the view name, and selects the
    ' matching view from the filtered "Recent List Views" results.
    Public Function SelectListView(ByVal sListViewName)
        Dim oPage, sLocViewOption
        Set oPage = GetAppPage()

        If Not WaitForObject(oPage.WebElement(sLocSelectListViewBtn), DEFAULT_SYNC_MS) Then
            ReportStep "PO_OpportunityList.SelectListView", False, "List View picker button not found within timeout"
            SelectListView = False
            Exit Function
        End If
        oPage.WebElement(sLocSelectListViewBtn).Click

        If Not WaitForObject(oPage.WebEdit(sLocSearchListsBox), DEFAULT_SYNC_MS) Then
            ReportStep "PO_OpportunityList.SelectListView", False, "Search lists box not found within timeout"
            SelectListView = False
            Exit Function
        End If
        oPage.WebEdit(sLocSearchListsBox).Click   ' ensure focus before Type - confirmed needed 2026-10-06
        oPage.WebEdit(sLocSearchListsBox).Type sListViewName

        sLocViewOption = "xpath:=//*[@role='option'][contains(.,'" & sListViewName & "')]"
        If Not WaitForObject(oPage.WebElement(sLocViewOption), DEFAULT_SYNC_MS) Then
            ReportStep "PO_OpportunityList.SelectListView", False, "List view option '" & sListViewName & "' not found within timeout"
            SelectListView = False
            Exit Function
        End If
        oPage.WebElement(sLocViewOption).Click
        SelectListView = True
    End Function

    ' Types into "Search this list..." and presses Enter to filter rows.
    Public Function SearchList(ByVal sSearchText)
        Dim oPage
        Set oPage = GetAppPage()
        If Not WaitForObject(oPage.WebEdit(sLocSearchThisListBox), DEFAULT_SYNC_MS) Then
            ReportStep "PO_OpportunityList.SearchList", False, "Search this list box not found within timeout"
            SearchList = False
            Exit Function
        End If
        oPage.WebEdit(sLocSearchThisListBox).Set sSearchText
        oPage.WebEdit(sLocSearchThisListBox).Type micReturn
        SearchList = True
    End Function

    ' Clicks the first Opportunity record link in the grid (identified by its
    ' href pattern "/lightning/r/006..." - 006 is the Opportunity id prefix).
    Public Function OpenFirstRecord()
        Dim oPage, sLocFirstRecord
        Set oPage = GetAppPage()
        sLocFirstRecord = "xpath:=(//a[contains(@href,'/lightning/r/006')])[1]"
        If WaitForObject(oPage.Link(sLocFirstRecord), DEFAULT_SYNC_MS) Then
            oPage.Link(sLocFirstRecord).Click
            OpenFirstRecord = True
        Else
            ReportStep "PO_OpportunityList.OpenFirstRecord", False, "First Opportunity record link not found within timeout"
            OpenFirstRecord = False
        End If
    End Function

End Class

' --- Page Object: Opportunity Stage Editor ------------------------------------
' NEW for this test case. Drives the inline "Edit Stage" panel directly
' (Details list pencil icon), NOT the Path's "Mark Stage as Complete" button
' used in Action_Opportunity.vbs. CONFIRMED locators captured via Playwright
' MCP on 2026-10-06.

Class PO_OpportunityStageEditor

    Private sLocEditStageBtn
    Private sLocStageDropdown
    Private sLocSubStageDropdown
    Private sLocSaveBtn

    Private Sub Class_Initialize()
        sLocEditStageBtn    = "xpath:=//button[@title='Edit Stage']"
        sLocStageDropdown   = "xpath:=//*[@role='combobox'][@aria-label='Stage']"
        sLocSubStageDropdown = "xpath:=//*[@role='combobox'][@aria-label='Sub-Stage']"
        sLocSaveBtn         = "xpath:=//button[text()='Save']"
    End Sub

    Public Function ClickEditStage()
        Dim oPage
        Set oPage = GetAppPage()
        If WaitForObject(oPage.WebButton(sLocEditStageBtn), DEFAULT_SYNC_MS) Then
            oPage.WebButton(sLocEditStageBtn).Click
            ClickEditStage = True
        Else
            ReportStep "PO_OpportunityStageEditor.ClickEditStage", False, "Edit Stage button not found within timeout"
            ClickEditStage = False
        End If
    End Function

    ' Opens the Stage dropdown and selects the given option (e.g. "Showround").
    Public Function SelectStage(ByVal sStageName)
        Dim oPage, sLocStageOption
        Set oPage = GetAppPage()

        If Not WaitForObject(oPage.WebElement(sLocStageDropdown), DEFAULT_SYNC_MS) Then
            ReportStep "PO_OpportunityStageEditor.SelectStage", False, "Stage dropdown not found within timeout"
            SelectStage = False
            Exit Function
        End If
        oPage.WebElement(sLocStageDropdown).Click

        sLocStageOption = "xpath:=//*[@role='option'][contains(.,'" & sStageName & "')]"
        If Not WaitForObject(oPage.WebElement(sLocStageOption), DEFAULT_SYNC_MS) Then
            ReportStep "PO_OpportunityStageEditor.SelectStage", False, "Stage option '" & sStageName & "' not found within timeout"
            SelectStage = False
            Exit Function
        End If
        oPage.WebElement(sLocStageOption).Click
        SelectStage = True
    End Function

    ' Opens the Sub-Stage dropdown and selects the given option
    ' (e.g. "Booked and Confirmed").
    Public Function SelectSubStage(ByVal sSubStageName)
        Dim oPage, sLocSubStageOption
        Set oPage = GetAppPage()

        If Not WaitForObject(oPage.WebElement(sLocSubStageDropdown), DEFAULT_SYNC_MS) Then
            ReportStep "PO_OpportunityStageEditor.SelectSubStage", False, "Sub-Stage dropdown not found within timeout"
            SelectSubStage = False
            Exit Function
        End If
        oPage.WebElement(sLocSubStageDropdown).Click

        sLocSubStageOption = "xpath:=//*[@role='option'][contains(.,'" & sSubStageName & "')]"
        If Not WaitForObject(oPage.WebElement(sLocSubStageOption), DEFAULT_SYNC_MS) Then
            ReportStep "PO_OpportunityStageEditor.SelectSubStage", False, "Sub-Stage option '" & sSubStageName & "' not found within timeout"
            SelectSubStage = False
            Exit Function
        End If
        oPage.WebElement(sLocSubStageOption).Click
        SelectSubStage = True
    End Function

    ' Clicks the Date textbox for a shared date/time field (e.g.
    ' "Site_Visit_Booked_Date_Time__c", "Assessment_Date_Time__c"), opens its
    ' date picker, and selects "Today". These fields render TWO inputs with
    ' the same name (Date + Time) - the Date one is the plain <input>, the
    ' Time one has id="combobox-input-*", so it's excluded explicitly.
    Public Function SetDateFieldToToday(ByVal sFieldName)
        Dim oPage, sLocDateInput, sLocDatePickerToggle, sLocToday
        Set oPage = GetAppPage()

        sLocDateInput = "xpath:=//input[@name='" & sFieldName & "'][not(contains(@id,'combobox'))]"
        If Not WaitForObject(oPage.WebEdit(sLocDateInput), DEFAULT_SYNC_MS) Then
            ReportStep "PO_OpportunityStageEditor.SetDateFieldToToday", False, "Date input for '" & sFieldName & "' not found within timeout"
            SetDateFieldToToday = False
            Exit Function
        End If
        oPage.WebEdit(sLocDateInput).Click

        sLocDatePickerToggle = "xpath:=(//input[@name='" & sFieldName & "'][not(contains(@id,'combobox'))]/following::button[contains(@title,'Select a date')])[1]"
        If Not WaitForObject(oPage.WebButton(sLocDatePickerToggle), DEFAULT_SYNC_MS) Then
            ReportStep "PO_OpportunityStageEditor.SetDateFieldToToday", False, "Date picker toggle for '" & sFieldName & "' not found within timeout"
            SetDateFieldToToday = False
            Exit Function
        End If
        oPage.WebButton(sLocDatePickerToggle).Click

        sLocToday = "xpath:=//button[text()='Today']"
        If Not WaitForObject(oPage.WebButton(sLocToday), DEFAULT_SYNC_MS) Then
            ReportStep "PO_OpportunityStageEditor.SetDateFieldToToday", False, "'Today' button not found within timeout"
            SetDateFieldToToday = False
            Exit Function
        End If
        oPage.WebButton(sLocToday).Click
        SetDateFieldToToday = True
    End Function

    ' Types into the "Showround feedback" textarea. This Lightning component
    ' has no stable id/name/aria-label - located relative to its label text
    ' (confirmed 2026-10-06: the label and textarea are siblings in document
    ' order, so "following::textarea[1]" reliably reaches it).
    Public Function SetShowroundFeedback(ByVal sText)
        Dim oPage, sLocFeedback
        Set oPage = GetAppPage()
        sLocFeedback = "xpath:=//*[text()='Showround feedback']/following::textarea[1]"
        If WaitForObject(oPage.WebEdit(sLocFeedback), DEFAULT_SYNC_MS) Then
            oPage.WebEdit(sLocFeedback).Set sText
            SetShowroundFeedback = True
        Else
            ReportStep "PO_OpportunityStageEditor.SetShowroundFeedback", False, "Showround feedback textarea not found within timeout"
            SetShowroundFeedback = False
        End If
    End Function

    ' Clicks a checkbox identified by its "name" attribute
    ' (e.g. "Assessment_Confirmed_by_Care_Home__c").
    Public Function ClickCheckboxByName(ByVal sFieldName)
        Dim oPage, sLocCheckbox
        Set oPage = GetAppPage()
        sLocCheckbox = "xpath:=//input[@name='" & sFieldName & "'][@type='checkbox']"
        If WaitForObject(oPage.WebElement(sLocCheckbox), DEFAULT_SYNC_MS) Then
            oPage.WebElement(sLocCheckbox).Click
            ClickCheckboxByName = True
        Else
            ReportStep "PO_OpportunityStageEditor.ClickCheckboxByName", False, "Checkbox '" & sFieldName & "' not found within timeout"
            ClickCheckboxByName = False
        End If
    End Function

    Public Function ClickSave()
        Dim oPage
        Set oPage = GetAppPage()
        If WaitForObject(oPage.WebButton(sLocSaveBtn), DEFAULT_SYNC_MS) Then
            oPage.WebButton(sLocSaveBtn).Click
            ClickSave = True
        Else
            ReportStep "PO_OpportunityStageEditor.ClickSave", False, "Save button not found within timeout"
            ClickSave = False
        End If
    End Function

End Class

' --- Page Object: Opportunity Related Tab / Field History ---------------------
' NEW for this test case. CONFIRMED locators captured via Playwright MCP on
' 2026-10-06.

Class PO_OpportunityRelated

    Private sLocRelatedTab
    Private sLocFieldHistoryFirstRow

    Private Sub Class_Initialize()
        sLocRelatedTab           = "xpath:=//*[@role='tab'][text()='Related']"
        sLocFieldHistoryFirstRow = "xpath:=//table[@aria-label='Opportunity Field History']/tbody/tr[1]"
    End Sub

    Public Function ClickRelatedTab()
        Dim oPage
        Set oPage = GetAppPage()
        If WaitForObject(oPage.WebElement(sLocRelatedTab), DEFAULT_SYNC_MS) Then
            oPage.WebElement(sLocRelatedTab).Click
            ClickRelatedTab = True
        Else
            ReportStep "PO_OpportunityRelated.ClickRelatedTab", False, "Related tab not found within timeout"
            ClickRelatedTab = False
        End If
    End Function

    ' Verifies the Opportunity Field History grid's FIRST row records a Stage
    ' change from sFromValue to sToValue. The table's cells render inside
    ' nested Lightning web components (shadow DOM), so raw xpath text
    ' matching on the cell itself doesn't work - instead we read the row's
    ' flattened "innertext" (which UFT's Web Add-in resolves through its own
    ' rendering, same as what the browser's accessibility tree exposes) and
    ' check it contains all three expected substrings.
    Public Function IsStageChangeRecorded(ByVal sFromValue, ByVal sToValue)
        Dim oPage, sRowText
        Set oPage = GetAppPage()

        If Not WaitForObject(oPage.WebElement(sLocFieldHistoryFirstRow), DEFAULT_SYNC_MS) Then
            ReportStep "PO_OpportunityRelated.IsStageChangeRecorded", False, "Opportunity Field History first row not found within timeout"
            IsStageChangeRecorded = False
            Exit Function
        End If

        sRowText = oPage.WebElement(sLocFieldHistoryFirstRow).GetROProperty("innertext")
        IsStageChangeRecorded = (InStr(sRowText, "Stage") > 0) And _
                                (InStr(sRowText, sFromValue) > 0) And _
                                (InStr(sRowText, sToValue) > 0)
    End Function

End Class

' --- Test flow -------------------------------------------------------------------

Dim oLogin, oNav, oOppList, oStageEditor, oRelated
Dim sUsername, sPassword, sVerificationCode
Dim bResult

' TODO: pull from an encrypted parameter / data table / environment variable -
' never hardcode real credentials in a checked-in script. Verification codes
' are typically one-time/short-lived - a real run will need a fresh one or a
' different verification method (e.g. trusted device/IP range) configured.
sUsername         = "qa_user@yourorg.com.sandboxname"
sPassword         = "<ENCRYPTED_PASSWORD>"
sVerificationCode = "<VERIFICATION_CODE>"

On Error Resume Next

' --- Step 1: Launch ----------------------------------------------------------
SystemUtil.Run APP_BROWSER & ".exe", APP_BROWSER_ARGS & " " & APP_URL
If Err.Number <> 0 Then
    ReportFatal "Launch Browser", "Failed to launch " & APP_BROWSER & ": " & Err.Description
End If

If Not WaitForObject(GetAppPage(), DEFAULT_PAGE_LOAD_MS) Then
    ReportFatal "Launch Browser", "Login page did not load within " & DEFAULT_PAGE_LOAD_MS & "ms"
End If

' --- Step 2: Login (username + password + verification code) -------------
Set oLogin = New PO_Login

If Not oLogin.IsLoginPageDisplayed() Then
    ReportFatal "Verify Login Page", "Login page not displayed"
End If
ReportStep "Verify Login Page", True, "Login page displayed as expected"

bResult = oLogin.Login(sUsername, sPassword, sVerificationCode)
ReportStep "Perform Login", bResult, "Login(username, password, verificationCode) invoked"

' --- Step 3: Open App Launcher, select Sales Centre app --------------------
Set oNav = New PO_AppNavigator

bResult = oNav.OpenAppLauncher()
ReportStep "Open App Launcher", bResult, "App Launcher button clicked"

bResult = oNav.OpenApp("Sales Centre")
ReportStep "Open Sales Centre App", bResult, "Searched and selected Sales Centre from App Launcher"

bResult = oNav.IsAppDisplayed("Sales Centre")
ReportStep "Verify Sales Centre App Displayed", bResult, "Sales Centre app heading displayed"

' --- Step 4: Navigate to Opportunities via Navigation Menu ------------------
bResult = oNav.ShowNavigationMenu()
ReportStep "Show Navigation Menu", bResult, "Navigation Menu opened"

bResult = oNav.ClickNavMenuItem("Opportunities")
ReportStep "Click Opportunities Menu Item", bResult, "Opportunities selected from Navigation Menu"

Set oOppList = New PO_OpportunityList

bResult = oOppList.IsOpportunitiesPageDisplayed()
ReportStep "Verify Opportunities Page Displayed", bResult, "Opportunities page heading displayed"

' --- Step 5: Select list view, search, open first record -------------------
bResult = oOppList.SelectListView("All Open Opportunities Care Enquiries")
ReportStep "Select List View", bResult, "Selected 'All Open Opportunities Care Enquiries' list view"

bResult = oOppList.SearchList("Assessment")
ReportStep "Search Opportunities List", bResult, "Searched list for 'Assessment'"

bResult = oOppList.OpenFirstRecord()
ReportStep "Open First Opportunity Record", bResult, "Opened first Opportunity record in filtered list"

' --- Step 6: First Stage cycle - Showround ----------------------------------
Set oStageEditor = New PO_OpportunityStageEditor

bResult = oStageEditor.ClickEditStage()
ReportStep "Click Edit Stage", bResult, "Edit Stage button clicked"

bResult = oStageEditor.SelectStage("Showround")
ReportStep "Select Showround Stage", bResult, "Showround selected from Stage dropdown"

bResult = oStageEditor.SelectSubStage("Booked and Confirmed")
ReportStep "Select Sub-Stage (Showround)", bResult, "Booked and Confirmed selected from Sub-Stage dropdown"

bResult = oStageEditor.SetDateFieldToToday("Site_Visit_Booked_Date_Time__c")
ReportStep "Set Site Visit Booked Date to Today", bResult, "Today selected from date picker"

bResult = oStageEditor.SetShowroundFeedback("Testing")
ReportStep "Set Showround Feedback", bResult, "'Testing' entered in Showround feedback"

bResult = oStageEditor.ClickSave()
ReportStep "Save Showround Stage", bResult, "Save button clicked"

' --- Step 7: Second Stage cycle - Assessment --------------------------------
bResult = oStageEditor.ClickEditStage()
ReportStep "Click Edit Stage (Assessment cycle)", bResult, "Edit Stage button clicked"

bResult = oStageEditor.SelectStage("Assessment")
ReportStep "Select Assessment Stage", bResult, "Assessment selected from Stage dropdown"

bResult = oStageEditor.SelectSubStage("Booked and Confirmed")
ReportStep "Select Sub-Stage (Assessment)", bResult, "Booked and Confirmed selected from Sub-Stage dropdown"

bResult = oStageEditor.ClickCheckboxByName("Assessment_Confirmed_by_Care_Home__c")
ReportStep "Check Assessment Confirmed by Care Home", bResult, "Checkbox clicked"

bResult = oStageEditor.SetDateFieldToToday("Assessment_Date_Time__c")
ReportStep "Set Assessment Date/Time to Today", bResult, "Today selected from date picker"

bResult = oStageEditor.ClickSave()
ReportStep "Save Assessment Stage", bResult, "Save button clicked"

' --- Step 8: Verify Stage change recorded in Field History ------------------
Set oRelated = New PO_OpportunityRelated

bResult = oRelated.ClickRelatedTab()
ReportStep "Click Related Tab", bResult, "Related tab clicked"

bResult = oRelated.IsStageChangeRecorded("Showround", "Assessment")
ReportStep "Verify Stage Change in Field History", bResult, "First Opportunity Field History row shows Stage changed from Showround to Assessment"

' --- Teardown --------------------------------------------------------------
Set oLogin = Nothing
Set oNav = Nothing
Set oOppList = Nothing
Set oStageEditor = Nothing
Set oRelated = Nothing
