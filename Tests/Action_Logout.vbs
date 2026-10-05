'==============================================================================
' Action_Logout.vbs
' Content for TC_001's Action2, renamed "Logout" and marked REUSABLE, so
' other test cases can end their session the same way instead of duplicating
' the logout flow.
'
' This file is NOT a Function Library - it is the literal content to paste
' into this Action's Editor/Expert View. See "UFT SETUP" below.
'
' UFT SETUP (one-time, inside the TC_001_Login_Logout test):
'   1. Insert > New Action (after "Login"), name it "Logout".
'   2. Action menu > Action Properties > check "This action is reusable or
'      can be called from other tests".
'   3. Action Properties > Parameters tab, add:
'         Output - LogoutStatus (Boolean)
'      (No inputs needed - it operates on whatever browser/page is already
'      open from the preceding Login action.)
'   4. This action relies on the SAME associated Function Libraries as
'      "Login" (Config, Utilities, PO_Login, PO_Home) - already associated
'      at the test level, no extra association needed here.
'   5. Open the "Logout" action's Editor/Expert View and paste everything
'      below this header block as its content.
'
' HOW OTHER TESTS CALL THIS (once this test is Associated via Insert > Call
' to Existing Action, pointing at this TC_001 test):
'
'   Dim outLogoutStatus
'   RunAction "Logout", oneIteration, outLogoutStatus
'   Reporter.ReportEvent IIf(outLogoutStatus, micPass, micFail), "Logout (via reusable action)", "Logout result"
'==============================================================================

Dim oLogin, oHome
Dim bResult

On Error Resume Next

' --- Step 1: Logout ---------------------------------------------------------
Set oHome = New PO_Home

bResult = oHome.Logout()
ReportStep "Perform Logout", bResult, "Logout() invoked via profile menu"

' --- Step 2: Verify back on Login page -------------------------------------
Set oLogin = New PO_Login

bResult = oLogin.IsLoginPageDisplayed()
ReportStep "Verify Logout", bResult, "Login page re-displayed after logout"

' --- Output parameter ------------------------------------------------------
Parameter("LogoutStatus") = bResult

' --- Teardown --------------------------------------------------------------
Set oLogin = Nothing
Set oHome = Nothing
GetAppBrowser().Close
