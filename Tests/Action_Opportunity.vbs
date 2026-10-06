' --- Config ------------------------------------------------------------------
Const SFQA_APP_URL               = "https://your-org--qa.sandbox.lightning.force.com/"  ' TODO: set to your actual sandbox/org URL
Const SFQA_APP_BROWSER           = "chrome"
Const SFQA_SYNC_TIMEOUT_MS       = 30000
Const SFQA_PAGE_LOAD_TIMEOUT_MS  = 60000
Const SFQA_BROWSER_TITLE_PATTERN = ".*Salesforce.*|.*Lightning.*"
Const SFQA_PAGE_TITLE_PATTERN    = ".*Salesforce.*|.*Lightning.*"

' --- Utilities -----------------------------------------------------------------

Function SFQA_GetAppBrowser()
    Dim oDesc
    Set oDesc = Description.Create()
    oDesc("title").Value = SFQA_BROWSER_TITLE_PATTERN
    oDesc("title").RegularExpression = True
    Set SFQA_GetAppBrowser = Browser(oDesc)
End Function

Function SFQA_GetAppPage()
    Dim oDesc
    Set oDesc = Description.Create()
    oDesc("title").Value = SFQA_PAGE_TITLE_PATTERN
    oDesc("title").RegularExpression = True
    Set SFQA_GetAppPage = SFQA_GetAppBrowser().Page(oDesc)
End Function

Function SFQA_WaitForObject(oTestObject, nTimeoutMs)
    Dim nTimeoutSec
    If nTimeoutMs = 0 Then nTimeoutMs = SFQA_SYNC_TIMEOUT_MS
    nTimeoutSec = nTimeoutMs / 1000
    SFQA_WaitForObject = oTestObject.Exist(nTimeoutSec)
End Function

Sub SFQA_ReportStep(sStepName, bPassed, sDetails)
    Dim nStatus
    If bPassed Then
        nStatus = micPass
    Else
        nStatus = micFail
    End If
    Reporter.ReportEvent nStatus, sStepName, sDetails
End Sub

Sub SFQA_ReportFatal(sStepName, sDetails)
    Reporter.ReportEvent micFail, sStepName, sDetails
    ExitAction
End Sub

' --- Page Object: Login (incl. two-factor Verification Code step) -------------
' CONFIRMED locators (captured via Playwright MCP on 2026-10-06).

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
        Set oPage = SFQA_GetAppPage()
        If SFQA_WaitForObject(oPage.WebEdit(sLocUsername), SFQA_SYNC_TIMEOUT_MS) Then
            oPage.WebEdit(sLocUsername).Set sUsername
            EnterUsername = True
        Else
            SFQA_ReportStep "PO_Login.EnterUsername", False, "Username field not found within timeout"
            EnterUsername = False
        End If
    End Function

    Public Function EnterPassword(ByVal sPassword)
        Dim oPage
        Set oPage = SFQA_GetAppPage()
        If SFQA_WaitForObject(oPage.WebEdit(sLocPassword), SFQA_SYNC_TIMEOUT_MS) Then
            oPage.WebEdit(sLocPassword).SetSecure sPassword
            EnterPassword = True
        Else
            SFQA_ReportStep "PO_Login.EnterPassword", False, "Password field not found within timeout (step 1 may have failed)"
            EnterPassword = False
        End If
    End Function

    Public Function ClickLogin()
        Dim oPage
        Set oPage = SFQA_GetAppPage()
        If SFQA_WaitForObject(oPage.WebButton(sLocBtnLogin), SFQA_SYNC_TIMEOUT_MS) Then
            oPage.WebButton(sLocBtnLogin).Click
            ClickLogin = True
        Else
            SFQA_ReportStep "PO_Login.ClickLogin", False, "Login button not found within timeout"
            ClickLogin = False
        End If
    End Function

    Public Function EnterVerificationCode(ByVal sCode)
        Dim oPage
        Set oPage = SFQA_GetAppPage()
        If SFQA_WaitForObject(oPage.WebEdit(sLocVerificationCode), SFQA_SYNC_TIMEOUT_MS) Then
            oPage.WebEdit(sLocVerificationCode).Set sCode
            EnterVerificationCode = True
        Else
            SFQA_ReportStep "PO_Login.EnterVerificationCode", False, "Verification Code field not found within timeout"
            EnterVerificationCode = False
        End If
    End Function

    Public Function ClickVerify()
        Dim oPage
        Set oPage = SFQA_GetAppPage()
        If SFQA_WaitForObject(oPage.WebButton(sLocBtnVerify), SFQA_SYNC_TIMEOUT_MS) Then
            oPage.WebButton(sLocBtnVerify).Click
            ClickVerify = True
        Else
            SFQA_ReportStep "PO_Login.ClickVerify", False, "Verify button not found within timeout"
            ClickVerify = False
        End If
    End Function

    ' Drives the full three-step flow: username -> Login (reveals password) ->
    ' password -> Login (reveals verification code) -> code -> Verify.
    Public Function Login(ByVal sUsername, ByVal sPassword, ByVal sVerificationCode)
        Dim bStep1, bStep2

        bStep1 = EnterUsername(sUsername) And ClickLogin()
        If Not bStep1 Then
            SFQA_ReportStep "PO_Login.Login", False, "Step 1 (username submit) failed"
            Login = False
            Exit Function
        End If

        bStep2 = EnterPassword(sPassword) And ClickLogin()
        If Not bStep2 Then
            SFQA_ReportStep "PO_Login.Login", False, "Step 2 (password submit) failed"
            Login = False
            Exit Function
        End If

        Login = EnterVerificationCode(sVerificationCode) And ClickVerify()
    End Function

    Public Function IsLoginPageDisplayed()
        Dim oPage
        Set oPage = SFQA_GetAppPage()
        IsLoginPageDisplayed = SFQA_WaitForObject(oPage.WebEdit(sLocUsername), SFQA_SYNC_TIMEOUT_MS)
    End Function

End Class

' --- Page Object: App Navigator (App Launcher + top Navigation Menu) ----------
' CONFIRMED locators captured via Playwright MCP on 2026-10-06.

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
        Set oPage = SFQA_GetAppPage()
        If SFQA_WaitForObject(oPage.WebButton(sLocAppLauncherBtn), SFQA_SYNC_TIMEOUT_MS) Then
            oPage.WebButton(sLocAppLauncherBtn).Click
            OpenAppLauncher = True
        Else
            SFQA_ReportStep "PO_AppNavigator.OpenAppLauncher", False, "App Launcher button not found within timeout"
            OpenAppLauncher = False
        End If
    End Function

    ' Types the app name into the App Launcher search box and clicks the
    ' matching app tile from the filtered results.
    Public Function OpenApp(ByVal sAppName)
        Dim oPage, sLocAppOption
        Set oPage = SFQA_GetAppPage()

        If Not SFQA_WaitForObject(oPage.WebEdit(sLocSearchAppsBox), SFQA_SYNC_TIMEOUT_MS) Then
            SFQA_ReportStep "PO_AppNavigator.OpenApp", False, "Search apps and items box not found within timeout"
            OpenApp = False
            Exit Function
        End If
        oPage.WebEdit(sLocSearchAppsBox).Set sAppName

        sLocAppOption = "xpath:=//*[@role='option'][contains(.,'" & sAppName & "')]"
        If Not SFQA_WaitForObject(oPage.WebElement(sLocAppOption), SFQA_SYNC_TIMEOUT_MS) Then
            SFQA_ReportStep "PO_AppNavigator.OpenApp", False, "App option '" & sAppName & "' not found within timeout"
            OpenApp = False
            Exit Function
        End If
        oPage.WebElement(sLocAppOption).Click
        OpenApp = True
    End Function

    Public Function ShowNavigationMenu()
        Dim oPage
        Set oPage = SFQA_GetAppPage()
        If SFQA_WaitForObject(oPage.WebButton(sLocShowNavMenuBtn), SFQA_SYNC_TIMEOUT_MS) Then
            oPage.WebButton(sLocShowNavMenuBtn).Click
            ShowNavigationMenu = True
        Else
            SFQA_ReportStep "PO_AppNavigator.ShowNavigationMenu", False, "Show Navigation Menu button not found within timeout"
            ShowNavigationMenu = False
        End If
    End Function

    ' Clicks a menu item (e.g. "Opportunities") from the open Navigation Menu.
    Public Function ClickNavMenuItem(ByVal sItemName)
        Dim oPage, sLocMenuItem
        Set oPage = SFQA_GetAppPage()
        sLocMenuItem = "xpath:=//*[@role='menuitem'][contains(.,'" & sItemName & "')]"
        If SFQA_WaitForObject(oPage.WebElement(sLocMenuItem), SFQA_SYNC_TIMEOUT_MS) Then
            oPage.WebElement(sLocMenuItem).Click
            ClickNavMenuItem = True
        Else
            SFQA_ReportStep "PO_AppNavigator.ClickNavMenuItem", False, "Menu item '" & sItemName & "' not found within timeout"
            ClickNavMenuItem = False
        End If
    End Function

    ' Verifies the app shell heading (top-left, e.g. "Sales Centre") is displayed.
    Public Function IsAppDisplayed(ByVal sAppName)
        Dim oPage, sLocHeading
        Set oPage = SFQA_GetAppPage()
        sLocHeading = "xpath:=//h1[text()='" & sAppName & "']"
        IsAppDisplayed = SFQA_WaitForObject(oPage.WebElement(sLocHeading), SFQA_PAGE_LOAD_TIMEOUT_MS)
    End Function

End Class

' --- Page Object: Opportunity List -----------------------------------------
' CONFIRMED locators captured via Playwright MCP on 2026-10-06.

Class PO_OpportunityList

    Private sLocSelectListViewBtn
    Private sLocSearchListsBox
    Private sLocSearchThisListBox

    Private Sub Class_Initialize()
        sLocSelectListViewBtn = "xpath:=//button[@title='Select a List View: Opportunities']"
        sLocSearchListsBox    = "xpath:=//input[@placeholder='Search lists...']"
        sLocSearchThisListBox = "xpath:=//input[@placeholder='Search this list...']"
    End Sub

    Public Function IsOpportunitiesPageDisplayed()
        Dim oPage, sLocHeading
        Set oPage = SFQA_GetAppPage()
        sLocHeading = "xpath:=//h1[text()='Opportunities']"
        IsOpportunitiesPageDisplayed = SFQA_WaitForObject(oPage.WebElement(sLocHeading), SFQA_PAGE_LOAD_TIMEOUT_MS)
    End Function

    ' Opens the list-view picker, types the view name, and selects the
    ' matching view from the filtered "Recent List Views" results.
    Public Function SelectListView(ByVal sListViewName)
        Dim oPage, sLocViewOption
        Set oPage = SFQA_GetAppPage()

        If Not SFQA_WaitForObject(oPage.WebButton(sLocSelectListViewBtn), SFQA_SYNC_TIMEOUT_MS) Then
            SFQA_ReportStep "PO_OpportunityList.SelectListView", False, "List View picker button not found within timeout"
            SelectListView = False
            Exit Function
        End If
        oPage.WebButton(sLocSelectListViewBtn).Click

        If Not SFQA_WaitForObject(oPage.WebEdit(sLocSearchListsBox), SFQA_SYNC_TIMEOUT_MS) Then
            SFQA_ReportStep "PO_OpportunityList.SelectListView", False, "Search lists box not found within timeout"
            SelectListView = False
            Exit Function
        End If
        oPage.WebEdit(sLocSearchListsBox).Set sListViewName

        sLocViewOption = "xpath:=//*[@role='option'][contains(.,'" & sListViewName & "')]"
        If Not SFQA_WaitForObject(oPage.WebElement(sLocViewOption), SFQA_SYNC_TIMEOUT_MS) Then
            SFQA_ReportStep "PO_OpportunityList.SelectListView", False, "List view option '" & sListViewName & "' not found within timeout"
            SelectListView = False
            Exit Function
        End If
        oPage.WebElement(sLocViewOption).Click
        SelectListView = True
    End Function

    ' Types into "Search this list..." and presses Enter to filter rows.
    Public Function SearchList(ByVal sSearchText)
        Dim oPage
        Set oPage = SFQA_GetAppPage()
        If Not SFQA_WaitForObject(oPage.WebEdit(sLocSearchThisListBox), SFQA_SYNC_TIMEOUT_MS) Then
            SFQA_ReportStep "PO_OpportunityList.SearchList", False, "Search this list box not found within timeout"
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
        Set oPage = SFQA_GetAppPage()
        sLocFirstRecord = "xpath:=(//a[contains(@href,'/lightning/r/006')])[1]"
        If SFQA_WaitForObject(oPage.Link(sLocFirstRecord), SFQA_SYNC_TIMEOUT_MS) Then
            oPage.Link(sLocFirstRecord).Click
            OpenFirstRecord = True
        Else
            SFQA_ReportStep "PO_OpportunityList.OpenFirstRecord", False, "First Opportunity record link not found within timeout"
            OpenFirstRecord = False
        End If
    End Function

End Class

' --- Page Object: Opportunity Detail ----------------------------------------
' CONFIRMED locators captured via Playwright MCP on 2026-10-06.

Class PO_OpportunityDetail

    Private sLocMarkStageCompleteBtn
    Private sLocStageDropdown
    Private sLocDoneBtn
    Private sLocStageChangedToast
    Private sLocEditMoveInDateBtn
    Private sLocMoveInDateTextbox
    Private sLocDatePickerTodayBtn
    Private sLocSaveBtn
    Private sLocClosedWonCard

    Private Sub Class_Initialize()
        sLocMarkStageCompleteBtn = "xpath:=(//button[contains(.,'Mark Stage as Complete')])[1]"
        sLocStageDropdown        = "xpath:=//*[@role='combobox'][@aria-label='Stage']"
        sLocDoneBtn              = "xpath:=//button[text()='Done']"
        sLocStageChangedToast    = "xpath:=//*[contains(text(),'Stage changed successfully')]"
        sLocEditMoveInDateBtn    = "xpath:=//button[@title='Edit Expected Move-in Date']"
        sLocMoveInDateTextbox    = "xpath:=//input[@name='Expected_Move_in_Date__c']"
        sLocDatePickerTodayBtn   = "xpath:=//button[text()='Today']"
        sLocSaveBtn              = "xpath:=//button[text()='Save']"
        sLocClosedWonCard        = "xpath:=//article[@class='slds-card'][contains(.,'Closed Won')]"
    End Sub

    Public Function ClickMarkStageAsComplete()
        Dim oPage
        Set oPage = SFQA_GetAppPage()
        If SFQA_WaitForObject(oPage.WebButton(sLocMarkStageCompleteBtn), SFQA_SYNC_TIMEOUT_MS) Then
            oPage.WebButton(sLocMarkStageCompleteBtn).Click
            ClickMarkStageAsComplete = True
        Else
            SFQA_ReportStep "PO_OpportunityDetail.ClickMarkStageAsComplete", False, "Mark Stage as Complete button not found within timeout"
            ClickMarkStageAsComplete = False
        End If
    End Function

    ' Opens the Stage dropdown (in the "Change Stage" dialog) and selects the
    ' given option (e.g. "Closed Won").
    Public Function ChangeStage(ByVal sStageName)
        Dim oPage, sLocStageOption
        Set oPage = SFQA_GetAppPage()

        If Not SFQA_WaitForObject(oPage.WebElement(sLocStageDropdown), SFQA_SYNC_TIMEOUT_MS) Then
            SFQA_ReportStep "PO_OpportunityDetail.ChangeStage", False, "Stage dropdown not found within timeout"
            ChangeStage = False
            Exit Function
        End If
        oPage.WebElement(sLocStageDropdown).Click

        sLocStageOption = "xpath:=//*[@role='option'][contains(.,'" & sStageName & "')]"
        If Not SFQA_WaitForObject(oPage.WebElement(sLocStageOption), SFQA_SYNC_TIMEOUT_MS) Then
            SFQA_ReportStep "PO_OpportunityDetail.ChangeStage", False, "Stage option '" & sStageName & "' not found within timeout"
            ChangeStage = False
            Exit Function
        End If
        oPage.WebElement(sLocStageOption).Click
        ChangeStage = True
    End Function

    Public Function ClickDone()
        Dim oPage
        Set oPage = SFQA_GetAppPage()
        If SFQA_WaitForObject(oPage.WebButton(sLocDoneBtn), SFQA_SYNC_TIMEOUT_MS) Then
            oPage.WebButton(sLocDoneBtn).Click
            ClickDone = True
        Else
            SFQA_ReportStep "PO_OpportunityDetail.ClickDone", False, "Done button not found within timeout"
            ClickDone = False
        End If
    End Function

    Public Function IsStageChangedMessageDisplayed()
        Dim oPage
        Set oPage = SFQA_GetAppPage()
        IsStageChangedMessageDisplayed = SFQA_WaitForObject(oPage.WebElement(sLocStageChangedToast), SFQA_SYNC_TIMEOUT_MS)
    End Function

    Public Function ClickEditMoveInDate()
        Dim oPage
        Set oPage = SFQA_GetAppPage()
        If SFQA_WaitForObject(oPage.WebButton(sLocEditMoveInDateBtn), SFQA_SYNC_TIMEOUT_MS) Then
            oPage.WebButton(sLocEditMoveInDateBtn).Click
            ClickEditMoveInDate = True
        Else
            SFQA_ReportStep "PO_OpportunityDetail.ClickEditMoveInDate", False, "Edit Expected Move-in Date button not found within timeout"
            ClickEditMoveInDate = False
        End If
    End Function

    ' Clicks into the Expected Move-in Date textbox, opens its date picker,
    ' and selects "Today".
    Public Function SetMoveInDateToToday()
        Dim oPage
        Set oPage = SFQA_GetAppPage()

        If Not SFQA_WaitForObject(oPage.WebEdit(sLocMoveInDateTextbox), SFQA_SYNC_TIMEOUT_MS) Then
            SFQA_ReportStep "PO_OpportunityDetail.SetMoveInDateToToday", False, "Expected Move-in Date textbox not found within timeout"
            SetMoveInDateToToday = False
            Exit Function
        End If
        oPage.WebEdit(sLocMoveInDateTextbox).Click

        If Not SFQA_WaitForObject(oPage.WebButton(sLocDatePickerTodayBtn), SFQA_SYNC_TIMEOUT_MS) Then
            SFQA_ReportStep "PO_OpportunityDetail.SetMoveInDateToToday", False, "Date picker 'Today' button not found within timeout"
            SetMoveInDateToToday = False
            Exit Function
        End If
        oPage.WebButton(sLocDatePickerTodayBtn).Click
        SetMoveInDateToToday = True
    End Function

    Public Function ClickSave()
        Dim oPage
        Set oPage = SFQA_GetAppPage()
        If SFQA_WaitForObject(oPage.WebButton(sLocSaveBtn), SFQA_SYNC_TIMEOUT_MS) Then
            oPage.WebButton(sLocSaveBtn).Click
            ClickSave = True
        Else
            SFQA_ReportStep "PO_OpportunityDetail.ClickSave", False, "Save button not found within timeout"
            ClickSave = False
        End If
    End Function

    Public Function IsClosedWonCardDisplayed()
        Dim oPage
        Set oPage = SFQA_GetAppPage()
        IsClosedWonCardDisplayed = SFQA_WaitForObject(oPage.WebElement(sLocClosedWonCard), SFQA_SYNC_TIMEOUT_MS)
    End Function

End Class

' --- Test flow -------------------------------------------------------------------

Dim oLogin, oNav, oOppList, oOppDetail
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
SystemUtil.Run SFQA_APP_BROWSER & ".exe", SFQA_APP_URL
If Err.Number <> 0 Then
    SFQA_ReportFatal "Launch Browser", "Failed to launch " & SFQA_APP_BROWSER & ": " & Err.Description
End If

If Not SFQA_WaitForObject(SFQA_GetAppPage(), SFQA_PAGE_LOAD_TIMEOUT_MS) Then
    SFQA_ReportFatal "Launch Browser", "Login page did not load within " & SFQA_PAGE_LOAD_TIMEOUT_MS & "ms"
End If

' --- Step 2: Login (username + password + verification code) -------------
Set oLogin = New PO_Login

If Not oLogin.IsLoginPageDisplayed() Then
    SFQA_ReportFatal "Verify Login Page", "Login page not displayed"
End If
SFQA_ReportStep "Verify Login Page", True, "Login page displayed as expected"

bResult = oLogin.Login(sUsername, sPassword, sVerificationCode)
SFQA_ReportStep "Perform Login", bResult, "Login(username, password, verificationCode) invoked"

' --- Step 3: Open App Launcher, select Sales Centre app --------------------
Set oNav = New PO_AppNavigator

bResult = oNav.OpenAppLauncher()
SFQA_ReportStep "Open App Launcher", bResult, "App Launcher button clicked"

bResult = oNav.OpenApp("Sales Centre")
SFQA_ReportStep "Open Sales Centre App", bResult, "Searched and selected Sales Centre from App Launcher"

bResult = oNav.IsAppDisplayed("Sales Centre")
SFQA_ReportStep "Verify Sales Centre App Displayed", bResult, "Sales Centre app heading displayed"

' --- Step 4: Navigate to Opportunities via Navigation Menu ------------------
bResult = oNav.ShowNavigationMenu()
SFQA_ReportStep "Show Navigation Menu", bResult, "Navigation Menu opened"

bResult = oNav.ClickNavMenuItem("Opportunities")
SFQA_ReportStep "Click Opportunities Menu Item", bResult, "Opportunities selected from Navigation Menu"

Set oOppList = New PO_OpportunityList

bResult = oOppList.IsOpportunitiesPageDisplayed()
SFQA_ReportStep "Verify Opportunities Page Displayed", bResult, "Opportunities page heading displayed"

' --- Step 5: Select list view, search, open first record -------------------
bResult = oOppList.SelectListView("All Open Opportunities Care Enquiries")
SFQA_ReportStep "Select List View", bResult, "Selected 'All Open Opportunities Care Enquiries' list view"

bResult = oOppList.SearchList("Assessment")
SFQA_ReportStep "Search Opportunities List", bResult, "Searched list for 'Assessment'"

bResult = oOppList.OpenFirstRecord()
SFQA_ReportStep "Open First Opportunity Record", bResult, "Opened first Opportunity record in filtered list"

' --- Step 6: Mark Stage as Complete -> Closed Won ---------------------------
Set oOppDetail = New PO_OpportunityDetail

bResult = oOppDetail.ClickMarkStageAsComplete()
SFQA_ReportStep "Click Mark Stage as Complete", bResult, "Mark Stage as Complete button clicked"

bResult = oOppDetail.ChangeStage("Closed Won")
SFQA_ReportStep "Select Closed Won Stage", bResult, "Closed Won selected from Stage dropdown"

bResult = oOppDetail.ClickDone()
SFQA_ReportStep "Click Done", bResult, "Done button clicked"

bResult = oOppDetail.IsStageChangedMessageDisplayed()
SFQA_ReportStep "Verify Stage Changed Message", bResult, "'Stage changed successfully.' toast displayed"

' --- Step 7: Edit Expected Move-in Date to Today, Save ----------------------
bResult = oOppDetail.ClickEditMoveInDate()
SFQA_ReportStep "Click Edit Expected Move-in Date", bResult, "Edit Expected Move-in Date button clicked"

bResult = oOppDetail.SetMoveInDateToToday()
SFQA_ReportStep "Set Move-in Date to Today", bResult, "Today selected from date picker"

bResult = oOppDetail.ClickSave()
SFQA_ReportStep "Click Save", bResult, "Save button clicked"

' --- Step 8: Verify Closed Won status on the record card -------------------
bResult = oOppDetail.IsClosedWonCardDisplayed()
SFQA_ReportStep "Verify Closed Won Status", bResult, "'Closed Won' displayed in article.slds-card"

' --- Teardown --------------------------------------------------------------
Set oLogin = Nothing
Set oNav = Nothing
Set oOppList = Nothing
Set oOppDetail = Nothing
