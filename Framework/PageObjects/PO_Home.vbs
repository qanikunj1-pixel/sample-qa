'==============================================================================
' PO_Home.vbs
' Page Object for the post-login Salesforce Home / App shell (Lightning).
' Associate as a Function Library AFTER Config.vbs and Utilities.vbs.
' Locators use plain string-based Descriptive Programming ("xpath:=...").
' TODO: all three below are still PLACEHOLDERS - we have not yet navigated
' past login to capture the real post-login DOM. Re-capture with Playwright
' MCP (or Object Spy) before trusting this class.
'==============================================================================

Class PO_Home

    Private sLocUserProfileIcon
    Private sLocLogoutMenuItem
    Private sLocGlobalHeader

    Private Sub Class_Initialize()
        ' TODO: Lightning Experience user avatar button, top-right global header.
        sLocUserProfileIcon = "xpath:=//a[contains(@class,'userProfile') or contains(@class,'avatar')]"

        ' TODO: "Log Out" item inside the profile flyout menu.
        sLocLogoutMenuItem = "xpath:=//a[text()='Log Out']"

        ' TODO: any stable element proving the Home/App shell rendered - used
        ' for post-login sync instead of guessing a fixed wait.
        sLocGlobalHeader = "xpath:=//div[contains(@class,'slds-global-header')]"
    End Sub

    ' --- Validations -----------------------------------------------------------

    Public Function IsHomePageDisplayed()
        Dim oPage
        Set oPage = GetAppPage()
        IsHomePageDisplayed = WaitForObject(oPage.WebElement(sLocGlobalHeader), DEFAULT_PAGE_LOAD_MS)
    End Function

    ' --- Actions -------------------------------------------------------------

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
