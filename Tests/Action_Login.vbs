'==============================================================================
' Action_Login.vbs
' Content for TC_001's Action1, renamed "Login" and marked REUSABLE, so other
' test cases (TC_002_Create_Lead, TC_003_Create_Account, ...) can call it via
' RunAction instead of duplicating the login flow.
'
' This file is NOT a Function Library - it is the literal content to paste
' into this Action's Editor/Expert View. See "UFT SETUP" below.
'
' UFT SETUP (one-time, inside the TC_001_Login_Logout test):
'   1. Rename Action1 to "Login" (right-click Action1 tab > Rename).
'   2. Action menu > Action Properties > check "This action is reusable or
'      can be called from other tests".
'   3. Action Properties > Parameters tab, add:
'         Input  - Username  (String)
'         Input  - Password  (String)   [Password type if your UFT version offers it]
'         Output - LoginStatus (Boolean)
'   4. File > Test Settings > Resources tab > Associate Function Library,
'      add in this exact order:
'         Framework\Config\Config.vbs
'         Framework\Common\Utilities.vbs
'         Framework\PageObjects\PO_Login.vbs
'         Framework\PageObjects\PO_Home.vbs
'   5. Open the "Login" action's Editor/Expert View and paste everything
'      below this header block as its content.
'
' HOW OTHER TESTS CALL THIS (once this test is Associated via Insert > Call
' to Existing Action, pointing at this TC_001 test):
'
'   Dim outLoginStatus
'   RunAction "Login", oneIteration, "qa_user@yourorg.com.sandboxname", "<ENCRYPTED_PASSWORD>", outLoginStatus
'   If Not outLoginStatus Then
'       Reporter.ReportEvent micFail, "Login (via reusable action)", "Login failed - aborting"
'       ExitTest
'   End If
'==============================================================================

Dim oLogin, oHome
Dim sUsername, sPassword
Dim bResult

sUsername = Parameter("Username")
sPassword = Parameter("Password")

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

' --- Output parameter ------------------------------------------------------
Parameter("LoginStatus") = bResult

Set oLogin = Nothing
Set oHome = Nothing
