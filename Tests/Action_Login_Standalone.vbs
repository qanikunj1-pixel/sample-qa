' --- Config ------------------------------------------------------------------
Const SFQA_APP_URL               = "https://your-org--qa.sandbox.lightning.force.com"  ' TODO: set to your actual sandbox/org URL
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

    Public Function Login(ByVal sUsername, ByVal sPassword)
        Dim bStep1
        bStep1 = EnterUsername(sUsername) And ClickLogin()
        If Not bStep1 Then
            SFQA_ReportStep "PO_Login.Login", False, "Step 1 (username submit) failed"
            Login = False
            Exit Function
        End If
        Login = EnterPassword(sPassword) And ClickLogin()
    End Function

    Public Function IsLoginPageDisplayed()
        Dim oPage
        Set oPage = SFQA_GetAppPage()
        IsLoginPageDisplayed = SFQA_WaitForObject(oPage.WebEdit(sLocUsername), SFQA_SYNC_TIMEOUT_MS)
    End Function

    Public Function IsLoginErrorDisplayed()
        Dim oPage
        Set oPage = SFQA_GetAppPage()
        IsLoginErrorDisplayed = SFQA_WaitForObject(oPage.WebElement(sLocLoginError), 5000)
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
        Set oPage = SFQA_GetAppPage()
        IsHomePageDisplayed = SFQA_WaitForObject(oPage.WebElement(sLocGlobalHeader), SFQA_PAGE_LOAD_TIMEOUT_MS)
    End Function

    Public Function Logout()
        Dim oPage
        Set oPage = SFQA_GetAppPage()

        If Not SFQA_WaitForObject(oPage.Link(sLocUserProfileIcon), SFQA_SYNC_TIMEOUT_MS) Then
            SFQA_ReportStep "PO_Home.Logout", False, "User profile icon not found within timeout"
            Logout = False
            Exit Function
        End If
        oPage.Link(sLocUserProfileIcon).Click

        If Not SFQA_WaitForObject(oPage.Link(sLocLogoutMenuItem), SFQA_SYNC_TIMEOUT_MS) Then
            SFQA_ReportStep "PO_Home.Logout", False, "Log Out menu item not found within timeout"
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
SystemUtil.Run SFQA_APP_BROWSER & ".exe", SFQA_APP_URL
If Err.Number <> 0 Then
    SFQA_ReportFatal "Launch Browser", "Failed to launch " & SFQA_APP_BROWSER & ": " & Err.Description
End If

If Not SFQA_WaitForObject(SFQA_GetAppPage(), SFQA_PAGE_LOAD_TIMEOUT_MS) Then
    SFQA_ReportFatal "Launch Browser", "Login page did not load within " & SFQA_PAGE_LOAD_TIMEOUT_MS & "ms"
End If

' --- Step 2: Login -------------------------------------------------------
Set oLogin = New PO_Login

If Not oLogin.IsLoginPageDisplayed() Then
    SFQA_ReportFatal "Verify Login Page", "Login page not displayed"
End If
SFQA_ReportStep "Verify Login Page", True, "Login page displayed as expected"

bResult = oLogin.Login(sUsername, sPassword)
SFQA_ReportStep "Perform Login", bResult, "Login(username, password) invoked"

' --- Step 3: Verify Home page --------------------------------------------
Set oHome = New PO_Home

bResult = oHome.IsHomePageDisplayed()
SFQA_ReportStep "Verify Home Page", bResult, "Home/App shell displayed after login"

If Not bResult Then
    SFQA_ReportFatal "Verify Home Page", "Home page did not render - aborting before logout step"
End If

' --- Step 4: Logout ----------------------------------------------------
bResult = oHome.Logout()
SFQA_ReportStep "Perform Logout", bResult, "Logout() invoked via profile menu"

' --- Step 5: Verify back on Login page -------------------------------------
bResult = oLogin.IsLoginPageDisplayed()
SFQA_ReportStep "Verify Logout", bResult, "Login page re-displayed after logout"

' --- Teardown --------------------------------------------------------------
Set oLogin = Nothing
Set oHome = Nothing
SFQA_GetAppBrowser().Close
