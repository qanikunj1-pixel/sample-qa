'==============================================================================
' TC_001_Login_Logout.vbs
' Test Case: TC_001_Salesforce_Login_Logout
'
' This is a Function Library (plain .vbs), same as the Page Objects - NOT a
' UFT Action itself. A UFT GUI Test's Action lives inside the test's own
' folder structure (Test.tsp / Action1\Script.mts / ...) and can't be a
' standalone .vbs file. So all real test logic stays here, version-controlled
' as text, and the Action in UFT becomes a ONE-LINE call into this Sub - see
' "UFT SETUP" below.
'
' Test steps:
'   1. Launch browser, navigate to Salesforce login.
'   2. Log in with valid credentials.
'   3. Verify Home/App shell is displayed.
'   4. Log out.
'   5. Verify return to Login page.
'==============================================================================

Sub TC_001_Login_Logout()

    Dim oLogin, oHome
    Dim sUsername, sPassword
    Dim bResult

    ' TODO: pull from an encrypted parameter / data table / environment variable -
    ' never hardcode real credentials in a checked-in script.
    sUsername = "qa_user@yourorg.com.sandboxname"
    sPassword = "<ENCRYPTED_PASSWORD>"

    On Error Resume Next

    ' --- Step 1: Launch ------------------------------------------------------
    SystemUtil.Run APP_BROWSER & ".exe", APP_URL
    If Err.Number <> 0 Then
        ReportFatal "Launch Browser", "Failed to launch " & APP_BROWSER & ": " & Err.Description
    End If

    If Not WaitForObject(GetAppPage(), DEFAULT_PAGE_LOAD_MS) Then
        ReportFatal "Launch Browser", "Login page did not load within " & DEFAULT_PAGE_LOAD_MS & "ms"
    End If

    ' --- Step 2: Login ---------------------------------------------------
    Set oLogin = New PO_Login

    If Not oLogin.IsLoginPageDisplayed() Then
        ReportFatal "Verify Login Page", "Login page not displayed"
    End If
    ReportStep "Verify Login Page", True, "Login page displayed as expected"

    bResult = oLogin.Login(sUsername, sPassword)
    ReportStep "Perform Login", bResult, "Login(username, password) invoked"

    ' --- Step 3: Verify Home page ------------------------------------------
    Set oHome = New PO_Home

    bResult = oHome.IsHomePageDisplayed()
    ReportStep "Verify Home Page", bResult, "Home/App shell displayed after login"

    If Not bResult Then
        ReportFatal "Verify Home Page", "Home page did not render - aborting before logout step"
    End If

    ' --- Step 4: Logout ----------------------------------------------------
    bResult = oHome.Logout()
    ReportStep "Perform Logout", bResult, "Logout() invoked via profile menu"

    ' --- Step 5: Verify back on Login page ----------------------------------
    bResult = oLogin.IsLoginPageDisplayed()
    ReportStep "Verify Logout", bResult, "Login page re-displayed after logout"

    ' --- Teardown ------------------------------------------------------------
    Set oLogin = Nothing
    Set oHome = Nothing
    GetAppBrowser().Close

End Sub
