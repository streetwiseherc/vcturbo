<%@ LANGUAGE = vbscript %>
<!--#INCLUDE FILE = "include/shop.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   PRODUCT_ALT.ASP                                                        #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>


<% Response.ExpiresAbsolute=#Jan 01, 1980 00:00:00# %>

<HTML>

<%
Dim rsProduct
Dim name
set rsProduct = Conn.Execute("select pfid, dept_id, manufacturer_id, name, short_description, image_filename, intro_date, date_changed, list_price, monogramable, long_description from vcturbo_product_family where pfid = '" & Replace(Request("pfid"), "'", "''") & "'")

if rsProduct.EOF then
    name = "Unknown"
else
    name = rsProduct("name").Value
end if
%>

<HEAD>
    <TITLE>Product '<% = mscsPage.HTMLEncode(name) %></TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>

<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>" BACKGROUND="<% = background %>">

<% hd_image = "heading_products.gif"%>
<% alt = "Product Page" %>
<!--#INCLUDE FILE = "include/header.asp" -->

<TABLE WIDTH="635">
    <% if rsProduct.EOF then %>
        <TR><TD>Product not available</TD></TR>
    <% else %>
        <TR>
            <TD colspan=3><H2><% = mscsPage.HTMLEncode(name) %></H2></TD>
        </TR>
        <TR>
            <TD valign=top>
                <% if not isnull(rsProduct("image_filename").Value) then %>
                    <IMG SRC="assets/prodimg/<% = rsProduct("image_filename").Value %>">
                <% else %>
                    Image not available
                <% end if %>
            </TD>
            <TD VALIGN=TOP WIDTH=400 colspan=2><% = mscsPage.HTMLEncode(rsProduct("short_description").Value) %></TD>
        </TR>
        <TR>
            <TD valign=top width=400 colspan=3><%= mscsPage.HTMLEncode(rsProduct("long_description").Value) %></TD>
        </TR>
        <TR>
            <FORM METHOD = "POST" ACTION="<% = pageURL("xt_orderform_edititem.asp") %>">
                <INPUT TYPE = HIDDEN NAME = "item_id" VALUE = "<% = CStr(Request("item_id")) %>" >
                <INPUT TYPE = HIDDEN NAME = "pfid"    VALUE = "<% = rsProduct("pfid").Value %>" >
                <TD><STRONG>Our Price:</STRONG>&nbsp;&nbsp;&nbsp;<FONT SIZE="+2"><%= MSCSDataFunctions.money(CLng(rsProduct("list_price").Value)) %></FONT></TD>
                <TD valign=top align=RIGHT><!-- #INCLUDE FILE = "include/prd_body.asp" --></TD>
                <TD valign=top align=LEFT>
                    <INPUT TYPE=image SRC="assets/images/btn_update.gif" BORDER=0 ALT="Update Basket" ALIGN=LEFT>
                </TD>
            </FORM>
        </TR>
    <% end if %>
</TABLE>
<% conn.close %>
<!--#INCLUDE FILE = "include/footer.asp" -->

</BODY>
</HTML>
