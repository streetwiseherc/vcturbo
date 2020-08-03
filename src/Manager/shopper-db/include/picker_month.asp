
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   PICKER.MONTH.ASP                                                        %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<SELECT NAME="IMONTH">
    <% = mscsPage.Option("01", imonth) %> Jan
    <% = mscsPage.Option("02", imonth) %> Feb
    <% = mscsPage.Option("03", imonth) %> Mar
    <% = mscsPage.Option("04", imonth) %> Apr
    <% = mscsPage.Option("05", imonth) %> May
    <% = mscsPage.Option("06", imonth) %> Jun
    <% = mscsPage.Option("07", imonth) %> Jul
    <% = mscsPage.Option("08", imonth) %> Aug
    <% = mscsPage.Option("09", imonth) %> Sep
    <% = mscsPage.Option("10", imonth) %> Oct
    <% = mscsPage.Option("11", imonth) %> Nov
    <% = mscsPage.Option("12", imonth) %> Dec
</SELECT>
