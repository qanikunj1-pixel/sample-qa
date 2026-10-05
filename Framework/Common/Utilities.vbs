'==============================================================================
' Utilities.vbs
' Associate as a Function Library in UFT One (loaded AFTER Config.vbs since it
' references Config constants). Generic, app-agnostic helpers shared by every
' Page Object and Driver script. No locators live here.
'==============================================================================

' --- Browser / Page descriptive handles --------------------------------------
' Every Page Object builds its controls under these two so the Browser/Page
' DP pattern is defined exactly once (Config.vbs), not re-typed per page.

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

' --- Synchronization -----------------------------------------------------------
' Wrap .Exist(timeout) instead of hard waits (Wait x). Returns True/False so
' calling code can decide to fail the step with a clear report entry.

Function WaitForObject(oTestObject, nTimeoutMs)
    Dim nTimeoutSec
    If nTimeoutMs = 0 Then nTimeoutMs = DEFAULT_SYNC_MS
    nTimeoutSec = nTimeoutMs / 1000
    WaitForObject = oTestObject.Exist(nTimeoutSec)
End Function

' --- Reporting -------------------------------------------------------------
' Single choke point for Reporter.ReportEvent so report formatting / future
' logging hooks (e.g. screenshot on fail) change in one place.

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
    ExitAction  ' or ExitTest, depending on how the driver is structured
End Sub
