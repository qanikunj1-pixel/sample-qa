'==============================================================================
' Driver_Login_Logout.vbs  -  UFT WIRING NOTES ONLY (not executable logic)
'
' The actual test logic lives in Framework\Tests\TC_001_Login_Logout.vbs as
' Sub TC_001_Login_Logout(). This file just documents how to wire a real UFT
' GUI Test to run it, since a UFT Action cannot itself be a plain .vbs file.
'
' HOW TO CREATE THE UFT TEST:
'   1. UFT One > File > New > Test > GUI Test. Name it e.g. "TC_001_Login_Logout".
'   2. File > Test Settings > Resources tab > Associate Function Library,
'      add these in this exact order (order matters - later files call
'      constants/functions defined in earlier ones):
'         Framework\Config\Config.vbs
'         Framework\Common\Utilities.vbs
'         Framework\PageObjects\PO_Login.vbs
'         Framework\PageObjects\PO_Home.vbs
'         Framework\Tests\TC_001_Login_Logout.vbs
'   3. Make sure the Web Add-in is loaded (Add-in Manager at UFT startup).
'   4. No Object Repository is used - do NOT associate a .tsr file.
'   5. Open Action1, switch to the Editor/Expert View, and replace its
'      contents with exactly ONE line:
'
'         TC_001_Login_Logout()
'
'   6. Save and run the test from UFT (or via command line / ALM).
'
' Why this split: every line of real logic stays in plain-text .vbs files
' that are diffable and version-controllable in this repo/IDE. The UFT test
' folder itself (Test.tsp, Action1\Script.mts, etc.) is just a thin shell
' that calls into it - there is nothing meaningful to review inside it.
'==============================================================================
