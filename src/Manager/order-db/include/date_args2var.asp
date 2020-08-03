
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   DATE_ARGS2VAR.ASP                                                       %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<%
if Request("imonth").Count = 0 then
    imonth = Month(Date)
else
    imonth = Cint(Request("imonth"))
end if

if imonth < 10 then
    imonth = "0" & imonth
else
    imonth = Cstr(imonth)
end if

if Request("iyear").Count = 0 then
    iyear = Cstr(Year(Date))
else
    iyear = Request("iyear")
end if
%>
