' --- Config ------------------------------------------------------------------
Const APP_URL               = "https://your-org--qa.sandbox.lightning.force.com"  ' TODO: set to your actual sandbox/org URL
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

' --- Page Object: Login --------------------------------------------------------

Class PO_Login

    Private sLocUsername
    Private sLocPassword
    Private sLocBtnLogin
    Private sLocLoginError

    Private Sub Class_Initialize()
        sLocUsername  = "xpath:=//input[@id='username']"
        sLocPassword  = "xpath:=//input[@id='password']"
        sLocBtnLogin  = "xpath:=//input[@id='Login']"
        sLocLoginError = "xpath:=//div[@id='error']"
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

    Public Function Login(ByVal sUsername, ByVal sPassword)
        Dim bStep1
        bStep1 = EnterUsername(sUsername) And ClickLogin()
        If Not bStep1 Then
            ReportStep "PO_Login.Login", False, "Step 1 (username submit) failed"
            Login = False
            Exit Function
        End If
        Login = EnterPassword(sPassword) And ClickLogin()
    End Function

    Public Function IsLoginPageDisplayed()
        Dim oPage
        Set oPage = GetAppPage()
        IsLoginPageDisplayed = WaitForObject(oPage.WebEdit(sLocUsername), DEFAULT_SYNC_MS)
    End Function

    Public Function IsLoginErrorDisplayed()
        Dim oPage
        Set oPage = GetAppPage()
        IsLoginErrorDisplayed = WaitForObject(oPage.WebElement(sLocLoginError), 5000)
    End Function

End Class

' --- Page Object: Home ----------------------------------------------------------

Class PO_Home

    Private sLocUserProfileIcon
    Private sLocLogoutMenuItem
    Private sLocGlobalHeader

    Private Sub Class_Initialize()
        sLocUserProfileIcon = "xpath:=//a[contains(@class,'userProfile') or contains(@class,'avatar')]"
        sLocLogoutMenuItem  = "xpath:=//a[text()='Log Out']"
        sLocGlobalHeader    = "xpath:=//div[contains(@class,'slds-global-header')]"
    End Sub

    Public Function IsHomePageDisplayed()
        Dim oPage
        Set oPage = GetAppPage()
        IsHomePageDisplayed = WaitForObject(oPage.WebElement(sLocGlobalHeader), DEFAULT_PAGE_LOAD_MS)
    End Function

    Public Function Logout()
        Dim oPage
        Set oPage = GetAppPage()

        If Not WaitForObject(oPage.Link(sLocUserProfileIcon), DEFAULT_SYNC_MS) Then
            ReportStep "PO_Home.Logout", False, "User profile icon not found within timeout"
            Logout = False
            Exit Function
        End If
        oPage.Link(sLocUserProfileIcon).Click

        If Not WaitForObject(oPage.Link(sLocLogoutMenuItem), DEFAULT_SYNC_MS) Then
            ReportStep "PO_Home.Logout", False, "Log Out menu item not found within timeout"
            Logout = False
            Exit Function
        End If
        oPage.Link(sLocLogoutMenuItem).Click

        Logout = True
    End Function

End Class

' --- Test flow -------------------------------------------------------------------

Dim oLogin, oHome
Dim sUsername, sPassword
Dim bResult

' TODO: pull from an encrypted parameter / data table / environment variable -
' never hardcode real credentials in a checked-in script.
sUsername = "qa_user@yourorg.com.sandboxname"
sPassword = "<ENCRYPTED_PASSWORD>"

On Error Resume Next

' --- Step 1: Launch ----------------------------------------------------------
SystemUtil.Run APP_BROWSER & ".exe", APP_BROWSER_ARGS & " " & APP_URL
If Err.Number <> 0 Then
    ReportFatal "Launch Browser", "Failed to launch " & APP_BROWSER & ": " & Err.Description
End If

If Not WaitForObject(GetAppPage(), DEFAULT_PAGE_LOAD_MS) Then
    ReportFatal "Launch Browser", "Login page did not load within " & DEFAULT_PAGE_LOAD_MS & "ms"
End If

' --- Step 2: Login -------------------------------------------------------
Set oLogin = New PO_Login

If Not oLogin.IsLoginPageDisplayed() Then
    ReportFatal "Verify Login Page", "Login page not displayed"
End If
ReportStep "Verify Login Page", True, "Login page displayed as expected"

bResult = oLogin.Login(sUsername, sPassword)
ReportStep "Perform Login", bResult, "Login(username, password) invoked"

' --- Step 3: Verify Home page --------------------------------------------
Set oHome = New PO_Home

bResult = oHome.IsHomePageDisplayed()
ReportStep "Verify Home Page", bResult, "Home/App shell displayed after login"

If Not bResult Then
    ReportFatal "Verify Home Page", "Home page did not render - aborting before logout step"
End If

' --- Step 4: Logout ----------------------------------------------------
bResult = oHome.Logout()
ReportStep "Perform Logout", bResult, "Logout() invoked via profile menu"

' --- Step 5: Verify back on Login page -------------------------------------
bResult = oLogin.IsLoginPageDisplayed()
ReportStep "Verify Logout", bResult, "Login page re-displayed after logout"

' --- Teardown --------------------------------------------------------------
Set oLogin = Nothing
Set oHome = Nothing
GetAppBrowser().Close
