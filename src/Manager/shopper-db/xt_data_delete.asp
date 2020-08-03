<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   XT_DATA_DELETE.ASP                                                      %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<%
Set errorList = Server.CreateObject("Commerce.SimpleList")
cmdTemp.CommandText = "SELECT * FROM vcturbo_shopper WHERE shopper_id = '" & Request.Form("shopper_id") & "'"

Set recordSet = Server.CreateObject("ADODB.Recordset")
recordSet.Open cmdTemp, , adOpenKeyset, adLockOptimistic

if recordSet.RecordCount <> 0 then
    
    recordSet.Delete
    Response.Redirect("shopper_list.asp")
else
	errorList.Add("This shopper does not exist in the shopper table.")
end if

if errorList.Count <> 0 then
%>
   <!--#INCLUDE FILE="include/error.asp" -->
<%
end if
%>
