' --- Config ------------------------------------------------------------------
Const APP_URL               = "https://your-org--qa.sandbox.lightning.force.com/"  ' TODO: set to your actual sandbox/org URL
Const APP_BROWSER           = "msedge"
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

' --- Page Object: Login --------------------------------------------------------

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
            ReportStep "PO_Login.EnterPassword", False, "Password field not found within timeout"
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

    ' Username -> Login -> Password -> Login -> Verification Code -> Verify
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

' --- Page Object: App Navigator -----------------------------------------------

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

    Public Function IsAppDisplayed(ByVal sAppName)
        Dim oPage, sLocHeading
        Set oPage = GetAppPage()
        sLocHeading = "xpath:=//h1[text()='" & sAppName & "']"
        IsAppDisplayed = WaitForObject(oPage.WebElement(sLocHeading), DEFAULT_PAGE_LOAD_MS)
    End Function

End Class

' --- Page Object: Opportunity List --------------------------------------------

Class PO_OpportunityList

    Private sLocSelectListViewBtn
    Private sLocSearchListsBox
    Private sLocSearchThisListBox

    Private Sub Class_Initialize()
        ' click target is the lightning-button-icon wrapper, not the inner button
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
        oPage.WebEdit(sLocSearchListsBox).Click   ' ensure focus before Type
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

    ' Date input shares its "name" with the Time input - exclude the
    ' combobox-input id to target the Date field specifically.
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

    ' No stable attribute on this field - located relative to its label text.
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

    ' Cells render in shadow DOM, so read the row's flattened innertext
    ' instead of xpath-matching the cell text directly.
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

' TODO: pull from an encrypted parameter / data table / environment variable
sUsername         = "qa_user@yourorg.com.sandboxname"
sPassword         = "<ENCRYPTED_PASSWORD>"
sVerificationCode = "<VERIFICATION_CODE>"

On Error Resume Next

' --- Step 1: Launch ----------------------------------------------------------
SystemUtil.Run APP_BROWSER & ".exe", APP_URL
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
