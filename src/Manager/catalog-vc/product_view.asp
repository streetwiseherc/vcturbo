<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   PRODUCT_VIEW.ASP                                                        %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "Product '" &  Request("pfid") & "'" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<%
cmdTemp.CommandText = "SELECT * FROM vcturbo_product_family where pfid = '" & Replace(Request("pfid"), "'", "''") & "'"
Set rsProduct = Server.CreateObject("ADODB.Recordset")
rsProduct.Open cmdTemp, , 1, 1
%>

<A HREF="<% = "product_edit.asp?" & mscsPage.URLArgs("pfid", Cstr(Request("pfid"))) %>"> <H2>Edit Product</H2> </A>

<TABLE>
    <TR>
      <TD VALIGN=TOP>ID:</TD>
        <TD VALIGN=TOP><% = rsProduct("pfid").Value %></TD>
    </TR>

    <TR>
      <TD VALIGN=TOP>Name:</TD>
        <TD VALIGN=TOP><% = rsProduct("name").Value %></TD>
    </TR>

    <TR>
      <TD VALIGN=TOP>Short Description:</TD>
        <TD VALIGN=TOP><% = rsProduct("short_description").Value %></TD>

    <TR>
      <TD VALIGN=TOP>Long Description:</TD>
        <TD VALIGN=TOP><% = rsProduct("long_description").Value %></TD>

    <TR>
      <TD VALIGN=TOP>Price:</TD>
        <TD VALIGN=TOP><% = MSCSDataFunctions.Money(Cstr(rsProduct("list_price").Value)) %></TD>
    </TR>

    <TR>
      <TD VALIGN=TOP>Dept:</TD>

        <% 
        cmdTemp.CommandText = "SELECT * FROM vcturbo_dept where dept_id = " & rsProduct("dept_id").Value
        Set rsDept = Server.CreateObject("ADODB.Recordset")
        rsDept.Open cmdTemp, , adOpenKeyset, adLockReadOnly
        %>
        <TD VALIGN=TOP><% = rsDept("name").Value %></TD>
    </TR>

</TABLE>

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
