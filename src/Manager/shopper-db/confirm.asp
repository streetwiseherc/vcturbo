<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   CONFIRM.ASP                                                             %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = mscsPage.RequestString("confirm_title", "Confirm Delete") %>
<% action = mscsPage.RequestString("confirm_action") %>
<% label = mscsPage.RequestString("confirm_label", "Delete") %>
<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<H1> <% = Request("confirm_message") %> </H1>

<FORM METHOD="POST" ACTION="<% = Request("confirm_action") %>">
    <% for each arg in Request.Form %>
        <% if Left(arg, 8) <> "confirm_" then %>
            <INPUT TYPE="HIDDEN" NAME="<% = arg %>" VALUE="<% = Request.Form(arg) %>">
        <% end if %>
    <% next %>
    <INPUT TYPE="SUBMIT" VALUE="<% = label %>">
    <INPUT TYPE="BUTTON" VALUE="Cancel" onClick="history.back()"> 
</FORM>

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->

