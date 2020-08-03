<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   DEFAULT.ASP                                                             %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>
<!--#Include File="include/SiteUtil.Asp"-->

<%
If Request("Reload").Count <> 0 then
    ReloadSite
    Response.redirect Request("URL")
End IF

If Request("Status").Count <> 0 then
    ToggleSiteStatus
    Response.redirect Request("URL")
End IF

GetStatus Status, RevStatus
PCFFiles = GetPCFFiles
%>

<frameset rows="100,*" FRAMEBORDER="0" FRAMESPACING="0">
	<frame src="navbar.asp" name="navbar" scrolling="No" noresize marginheight="0" marginwidth="0" framespacing="0" frameborder="No">
	<frame src="manager.asp" name="body" FRAMEBORDER="0" FRAMESPACING="0">
</frameset>
