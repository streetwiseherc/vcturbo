<%@ LANGUAGE=vbscript  %>
<!--#INCLUDE FILE = "include/shop.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   LISTING.ASP                                                            #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>

<HTML>

<HEAD>
    <TITLE>Product Listing Page </TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>

<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>" BACKGROUND="<% = background %>">

<% hd_image = "heading_products.gif "%>
<% alt = "Product Listing Page" %>
<!--#INCLUDE FILE = "include/header.asp" -->

<% 
Dim rsProduct
Dim currenttopdept
Dim currentdept
Dim rowidx

REM Create Recordset, execute Query
Set cmdTemp = Server.CreateObject("ADODB.Command")
cmdTemp.CommandType = adCmdText
Set cmdTemp.ActiveConnection = conn
cmdTemp.CommandText = "SELECT d1.name AS top_name, d2.name AS dname, p.name AS pname, d1.dept_id AS dd1, d2.dept_id AS dd2, p.pfid FROM vcturbo_product_family p, vcturbo_dept d1, vcturbo_dept d2 WHERE ((d1.dept_id=d2.parent_id) AND (p.dept_id=d2.dept_id)) ORDER BY d1.dept_id, d2.dept_id"
Set rsProduct = Server.CreateObject("ADODB.Recordset")
rsProduct.Open cmdTemp, , adOpenStatic, adLockReadOnly
%>

<%

Dim fldPFID    
Dim fldTopName 
Dim fldDD1     
Dim fldDD2     
Dim fldDname   
Dim fldPname   

Set fldPFID    = rsProduct("pfid")
Set fldTopName = rsProduct("top_name")
Set fldDD1     = rsProduct("dd1")
Set fldDD2     = rsProduct("dd2")
Set fldDname   = rsProduct("dname")
Set fldPname   = rsProduct("pname")
%>

<TABLE WIDTH=700 BORDER=0>
    <TR>
        </TD>
        <% Do until rsProduct.EOF %>     
            </TD>
            <TD VALIGN=TOP WIDTH=225>
            <FONT SIZE=5><B><% = mscsPage.HTMLEncode(fldTopName.Value) %></B></FONT><BR> 
            <% currenttopdept = fldDD1.Value %>
            <% Do %>
                <% currentdept = fldDD2.Value %>
                <% rowidx = 1 %>
                <BR>&nbsp;
                <FONT SIZE=4><% = mscsPage.HTMLEncode(fldDname.Value) %></FONT><BR>
                    <% Do %>
                        &nbsp;&nbsp;
                        <A HREF = "<% = baseURL("product.asp") & mscsPage.URLShopperArgs("pfid", fldPFID.Value)%>"><% = mscsPage.HTMLEncode(fldPname.Value) %></A><br>
                        <% rowidx = rowidx + 1 %>
                        <% rsProduct.MoveNext %>
                        <% If rsProduct.EOF Then Exit Do %>
                    <% Loop While fldDD2 = currentdept %>
                    <% If rsProduct.EOF Then Exit Do %>
            <% Loop While fldDD1.Value = currenttopdept %>
            <BR>
            <BR>
        <% Loop %>
    </TR>
</TABLE>
<% conn.close %>
<!--#INCLUDE FILE = "include/footer.asp" -->

</BODY>
</HTML>
