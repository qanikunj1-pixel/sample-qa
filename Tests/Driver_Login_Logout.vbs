'==============================================================================
' Driver_Login_Logout.vbs
' Test Case: TC_001_Salesforce_Login_Logout
'
' This file is NOT a Function Library. It is the literal content to paste
' into Action1's Editor/Expert View in the UFT GUI Test. Keeping it here as
' plain text lets you diff/review/version-control the test steps even though
' UFT itself stores Action1's real copy inside the test's own folder
' structure (Test.tsp / Action1\Script.mts).
'
' UFT SETUP (one-time):
'   1. UFT One > File > New > Test > GUI Test. Name it "TC_001_Login_Logout".
'   2. File > Test Settings > Resources tab > Associate Function Library,
'      add in this exact order:
'         Framework\Config\Config.vbs
'         Framework\Common\Utilities.vbs
'         Framework\PageObjects\PO_Login.vbs
'         Framework\PageObjects\PO_Home.vbs
'   3. Make sure the Web Add-in is loaded (Add-in Manager at UFT startup).
'   4. No Object Repository is used - do NOT associate a .tsr file.
'   5. Open Action1, switch to the Editor/Expert View, and paste everything
'      below this header block as Action1's content.
'   6. Save and run.
'
' Whenever this test's steps change, edit BOTH this file (for review/history)
' AND Action1 inside the actual UFT test (for execution) - they are two
' copies of the same logic by design, since UFT can't execute this file
' directly.
'==============================================================================

Dim oLogin, oHome
Dim sUsername, sPassword
Dim bResult

' TODO: pull from an encrypted parameter / data table / environment variable -
' never hardcode real credentials in a checked-in script.
sUsername = "qa_user@yourorg.com.sandboxname"
sPassword = "<ENCRYPTED_PASSWORD>"

On Error Resume Next

' --- Step 1: Launch ----------------------------------------------------------
SystemUtil.Run APP_BROWSER & ".exe", APP_URL
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

' --- Step 4: Logout --------------------------------------------------------
bResult = oHome.Logout()
ReportStep "Perform Logout", bResult, "Logout() invoked via profile menu"

' --- Step 5: Verify back on Login page -------------------------------------
bResult = oLogin.IsLoginPageDisplayed()
ReportStep "Verify Logout", bResult, "Login page re-displayed after logout"

' --- Teardown --------------------------------------------------------------
Set oLogin = Nothing
Set oHome = Nothing
GetAppBrowser().Close
