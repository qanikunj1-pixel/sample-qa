'==============================================================================
' PO_Login.vbs
' Page Object for the Salesforce Login page.
' Associate as a Function Library AFTER Config.vbs and Utilities.vbs.
' No Object Repository - every control is identified with plain STRING-based
' Descriptive Programming ("property:=value"), since each one only needs a
' single xpath property. Description.Create() objects are only worth the
' extra code when a locator needs MULTIPLE properties combined.
'
' CONFIRMED locators (captured via Playwright MCP against the live sandbox
' login page on 2026-10-05). This is a TWO-STEP login:
'   Step 1: enter Username, click "Log In to Sandbox"  -> reveals Password field
'   Step 2: enter Password, click "Log In to Sandbox" again (SAME button xpath)
'==============================================================================

Class PO_Login

    Private sLocUsername
    Private sLocPassword
    Private sLocBtnLogin
    Private sLocLoginError

    Private Sub Class_Initialize()
        ' CONFIRMED: <input id="username" name="username" type="email">
        sLocUsername = "xpath:=//input[@id='username']"

        ' CONFIRMED: <input id="password" name="pw" type="password">
        ' Only rendered AFTER step 1's Login click, so always sync on it
        ' before interacting (see EnterPassword).
        sLocPassword = "xpath:=//input[@id='password']"

        ' CONFIRMED: <input id="Login" name="Login" type="submit" value="Log In to Sandbox">
        ' Same element/id is reused for BOTH step 1 and step 2 submits.
        sLocBtnLogin = "xpath:=//input[@id='Login']"

        ' NOT YET CONFIRMED - inferred only from username field's
        ' aria-describedby="error" attribute; no failed-login scenario was
        ' run yet. Re-verify on an invalid credentials attempt before relying
        ' on IsLoginErrorDisplayed.
        sLocLoginError = "xpath:=//div[@id='error']"
    End Sub

    ' --- Actions -------------------------------------------------------------

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
        ' Password field only exists after step 1's Login click - sync here
        ' doubles as the "did step 1 succeed" check.
        If WaitForObject(oPage.WebEdit(sLocPassword), DEFAULT_SYNC_MS) Then
            oPage.WebEdit(sLocPassword).SetSecure sPassword   ' use SetSecure for encoded passwords
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

    ' --- Convenience / composite action ---------------------------------------
    ' Drives the full two-step flow: username -> Login (reveals password) ->
    ' password -> Login (final submit).

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

    ' --- Validations -----------------------------------------------------------

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
