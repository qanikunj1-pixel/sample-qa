'==============================================================================
' Config.vbs
' Associate as a Function Library in UFT One (Resources > Associate Function
' Library). Holds environment/app level constants only - no locators here.
'==============================================================================

Const APP_URL           = "https://your-org--qa.sandbox.lightning.force.com"  ' TODO: set to your actual sandbox/org URL (redirects to .my.salesforce.com login)
Const APP_BROWSER       = "msedge"                           ' msedge | chrome | iexplore | firefox
Const APP_BROWSER_ARGS  = "-inprivate"                       ' Edge/Chromium private window switch
Const DEFAULT_SYNC_MS   = 30000                              ' default object sync timeout (ms)
Const DEFAULT_PAGE_LOAD_MS = 60000

' Salesforce DOM shell changes between Lightning/Classic and release versions.
' Keep the browser/page title patterns here so every Page Object references
' the SAME descriptive pattern instead of each one hardcoding it.
Const BROWSER_TITLE_PATTERN = ".*Salesforce.*|.*Lightning.*"
Const PAGE_TITLE_PATTERN    = ".*Salesforce.*|.*Lightning.*"
