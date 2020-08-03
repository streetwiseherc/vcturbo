
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   PICKER.YEAR.ASP                                                         %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<SELECT NAME="IYEAR">
    <% = mscsPage.Option("1998", iyear) %> 1998
    <% = mscsPage.Option("1999", iyear) %> 1999
    <% = mscsPage.Option("2000", iyear) %> 2000
    <% = mscsPage.Option("2001", iyear) %> 2001
    <% = mscsPage.Option("2002", iyear) %> 2002
    <% = mscsPage.Option("2003", iyear) %> 2003
</SELECT>
