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
Select Case Request.Form("type") 
    case "product2"
        cmdTemp.CommandText = "SELECT * FROM vcturbo_product_family WHERE pfid = '" & Replace(Request.Form("pfid"), "'", "''") & "'"
    Case "product" 
        cmdTemp.CommandText = "SELECT * FROM vcturbo_product_family WHERE pfid = '" & Replace(Request.Form("pfid"), "'", "''") & "'"
    Case "variant"
        cmdTemp.CommandText = "SELECT * FROM vcturbo_product_variant WHERE sku = " & Request.Form("sku")
    Case "department"
        cmdTemp.CommandText = "SELECT * FROM vcturbo_dept WHERE dept_id = " & Request.Form("dept_id")
End Select

Set recordSet = Server.CreateObject("ADODB.Recordset")
recordSet.Open cmdTemp, , adOpenKeyset, adLockOptimistic

if recordSet.RecordCount <>0 then

    if Request.Form("type") <> "product2" then
        recordSet.Delete
    end if

    Select Case Request.Form("type")
        case "product2"
            conn.Execute "delete from vcturbo_product_variant where pfid = '" & Replace(Request.Form("pfid"), "'", "''") & "'", , adCmdText
            conn.Execute "delete from vcturbo_product_attribute where pfid = '" & Replace(Request.Form("pfid"), "'", "''") & "'", , adCmdText
        Case "product"
            conn.Execute "delete from vcturbo_product_variant where pfid = '" & Replace(Request.Form("pfid"), "'", "''") & "'", , adCmdText
            conn.Execute "delete from vcturbo_product_attribute where pfid = '" & Replace(Request.Form("pfid"), "'", "''") & "'", , adCmdText
    End Select

    If Request.Form("pfid").Count = 0 Then
        Response.redirect Cstr(Request.Form("goto"))
    else
        If Request.Form("attribute_id").Count = 0 Then
            Response.redirect Cstr(Request.Form("goto")) & "?" & mscsPage.URLArgs("pfid", CStr(Request.Form("pfid")))
        else
            Response.redirect Cstr(Request.Form("goto")) & "?" & mscsPage.URLArgs("pfid", CStr(Request.Form("pfid")), "attribute_id", CStr(Request.Form("attribute_id")))
        end if
    End If
else
    Select Case Request.Form("type") 
        Case "product" 
            errorList.Add("This pfid does not exist in the product family table.")
        Case "variant"
            errorList.Add("This sku does not exist in the product variant table.")
        Case "department"
            errorList.Add("This department id does not exist in the department table.")
    End Select
end if

if errorList.Count <> 0 then
%>
        <!--#INCLUDE FILE="include/error.asp" -->

<%
end if
%>
