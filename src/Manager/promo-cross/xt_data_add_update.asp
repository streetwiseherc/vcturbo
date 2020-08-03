<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   XT_DATA_ADD_UPDATE.ASP                                                  %>
<% REM   VC TURBO STORE                                                          %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<%
Set errorList = Server.CreateObject("Commerce.SimpleList")
pfid = mscsPage.RequestString("pfid", null, 1, 30)
if IsNull(pfid) then
    errorList.Add("Product must be between 1 and 30 characters.")
end if      
related_pfid = mscsPage.RequestString("related_pfid",null,1,30)
if IsNull(related_pfid) then
    errorList.Add("Related product must be a between 1 and 30 characters")
end if
description = mscsPage.RequestString("description")
if Not(IsNull(description)) and len(description) > 255 then
    errorList.Add("Slogan must be less than 255 characters.")
end if
if pfid = related_pfid then
    errorList.Add("Cross promo product and related product must be different.")
end if

if errorList.count = 0 then
	mscsPromoCross = "Select pfid, related_pfid, description from vcturbo_promo_cross where pfid = ':1' and related_pfid = ':2'"
	sql = Replace(mscsPromoCross, "':1'", "'" & Replace(pfid, "'", "''") & "'")
	sql = Replace(sql, "':2'", "'" & Replace(related_pfid, "'", "''") & "'")
    cmdTemp.CommandText = sql
    
    Set recordSet = Server.CreateObject("ADODB.Recordset")
    recordSet.Open cmdTemp, , adOpenKeyset, adLockOptimistic

    op = mscsPage.RequestString("op", "add")

    if (op = "add" and recordSet.RecordCount = 0) or (op = "update" and recordSet.RecordCount = 1) then

        if op = "add" then
            recordSet.AddNew
        end if

        if op = "add" then
             recordSet("related_pfid").Value = related_pfid
             recordSet("pfid").Value  = pfid
        end if
        recordSet("description").Value = description

        recordSet.Update

        Response.Redirect "promo-cross_list.asp"

    else
        if op = "add" then
            errorList.Add("This cross promo already exists in the cross sell table.")
        else
            errorList.Add("This cross promo does not exist in the cross sell table.")
        end if

    end if
end if
if errorList.Count <> 0 then
%>
    <!--#INCLUDE FILE="include/error.asp" -->
<%
end if
%>
 
