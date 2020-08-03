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
dim errorList
dim sql
Set errorList = Server.CreateObject("Commerce.SimpleList")

sql = "SELECT * FROM vcturbo_promo_cross WHERE pfid = '" & Replace(Request.Form("pfid"), "'", "''") & "' and related_pfid = '" & Replace(Request.Form("rpfid"), "'", "''") & "'"
cmdTemp.CommandText = sql
Set recordSet = Server.CreateObject("ADODB.Recordset")
recordSet.Open cmdTemp, , adOpenKeyset, adLockOptimistic

if recordSet.RecordCount <>0 then
    
    recordSet.Delete

	Response.Redirect "promo-cross_list.asp"
else
	errorList.Add("This cross sell promotion does not exist in the cross sell table.")
end if

if errorList.Count <> 0 then
%>
        <!--#INCLUDE FILE="include/error.asp" -->
<%
end if
%>
