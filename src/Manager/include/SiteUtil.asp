<% 
REM ######################################################################
REM                                                                      #
REM   Microsoft Commerce Server v3.00                                    #
REM                                                                      #
REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved. #
REM                                                                      #
REM ######################################################################

'========================================================================
' ReportOnError
'
Sub ReportOnError(ByVal strErr)
    If Err.Number <> 0 Then
        Response.Write "<H3>" & strErr & "</H3>"
        Response.Write "<BLOCKQUOTE>" & Replace(Err.Description, "~", _
                                                vbCr + "<BR>&nbsp;&nbsp;")
        Response.Write Hex(Err.Number) & "</BLOCKQUOTE>" & vbCr
        Response.End
    End If
End Sub


'========================================================================
' GetSiteObject
'
Function GetSiteObject
    On Error Resume Next 

    Set SiteObj = Server.CreateObject("Commerce.AdminSite")
    ReportOnError "Unable to create object.  Check Commerce installation."
    
    SiteObj.InitializeFromMDPath Request.ServerVariables("APPL_MD_PATH")
    Set GetSiteObject = SiteObj
End Function


'========================================================================
' OpenSite
'
Sub OpenSite
    On Error Resume Next 

    Dim SiteObj
    Set SiteObj = GetSiteObject
    SiteObj.Status = TRUE

    ReportOnError "Unable to open site."
End Sub


'========================================================================
' CloseSite
'
Sub CloseSite
    On Error Resume Next 

    Dim SiteObj
    Set SiteObj = GetSiteObject
    SiteObj.Status = FALSE

    ReportOnError "Unable to close site."
End Sub


'========================================================================
' ToggleSiteStatus
'
Sub ToggleSiteStatus
    On Error Resume Next 

    Dim SiteObj
    Set SiteObj = GetSiteObject
    SiteObj.Status = Not Siteobj.Status

    ReportOnError "Unable to change site status."
End Sub


'========================================================================
' GetStatus
'
Sub GetStatus(ByRef Status, ByRef RevStatus)
    On Error Resume Next

    Dim SiteObj
    Set SiteObj = GetSiteObject

    Status    = "Invalid"
    RevStatus = "Invalid"

    Dim IsOpen
    IsOpen = SiteObj.Status

    If Err.Number = 0 Then
        If IsOpen Then 
            Status    = "Open"
            RevStatus = "Close"
        Else
            Status    = "Closed"
            RevStatus = "Open"
        End If
    End If
    Err.Clear
End Sub


'========================================================================
' ReloadSite
'
Sub ReloadSite
    On Error Resume Next 

    Dim SiteObj
    Set SiteObj = GetSiteObject
    SiteObj.Reload

    ReportOnError "Unable to reload site."
End Sub


'========================================================================
' GetPCFFiles
'
Function GetPCFFiles
    Dim AdminFiles
    Set AdminFiles = Server.CreateObject("Commerce.AdminFiles")

    Dim ConfigDir
    ConfigDir = Request.ServerVariables("APPL_PHYSICAL_PATH") + "\Config"

    GetPCFFiles = AdminFiles.GetFiles(ConfigDir + "\*.pcf")
End Function

%>
