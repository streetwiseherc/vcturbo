<%@ LANGUAGE = vbscript %>
<!--#INCLUDE FILE = "include/shop.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   PRODUCT_EDIT.ASP                                                       #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>
<% Response.ExpiresAbsolute=#Jan 01, 1980 00:00:00# %>

<%
Dim rsProduct
set rsProduct = Conn.Execute("select pfid, dept_id, manufacturer_id, name, short_description, image_filename, intro_date, date_changed, list_price, monogramable, long_description from vcturbo_product_family where pfid = '" & Replace(Request("pfid"), "'", "''") & "'")

Dim name
if rsProduct.EOF then
    name = "Unknown product"
else
    name = rsProduct("name").Value
end if
%>

<HTML>

<HEAD>
    <TITLE>Product '<% = mscsPage.HTMLEncode(name) %></TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>

<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>" BACKGROUND="<% = background %>">

<% hd_image = "heading_products.gif" %>
<% alt = "Product Page" %>
<!--#INCLUDE FILE = "include/header.asp" -->


<TABLE WIDTH="635">
    <% if rsProduct.EOF then %>
        <TR><TD>Product not available</TD></TR>
    <% else %>
        <TR>
            <TD COLSPAN=3><H2><% = mscsPage.HTMLEncode(name) %></H2></TD>
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
           <TD valign=top width=400 colspan=3><% = mscsPage.HTMLEncode(rsProduct("long_description").Value) %></TD>
        </TR>
        <TR>
        <FORM METHOD = "POST" ACTION="<% = pageURL("xt_orderform_edititem.asp") %>">
            <INPUT TYPE = HIDDEN NAME = "item_id" VALUE = "<% = CStr(Request("item_id")) %>" >
            <INPUT TYPE = HIDDEN NAME = "pfid" VALUE = "<% = rsProduct("pfid").Value %>" >
            <TD VALIGN=top><STRONG>Our Price:</STRONG>&nbsp;<FONT SIZE="+2"><% =MSCSDataFunctions.money(CLng(rsProduct("list_price"))) %></FONT></TD>
            <TD VALIGN=top ALIGN=RIGHT><!-- #INCLUDE FILE="include/prd_body.asp" --></TD>
            <TD VALIGN=TOP ALIGN=LEFT>
            <TABLE>
                <TR>
                    <TD><INPUT TYPE=image SRC="assets/images/btn_update.gif" BORDER=0 ALT="Update Basket" ALIGN=LEFT></TD>
                </TR>
                <TR>
                    <TD>
                        <A HREF="<% = baseURL("xt_orderform_delitem.asp") & mscsPage.URLShopperArgs("item_id",  CStr(Request("item_id"))) %>">
                            <IMG SRC="assets/images/btn_remove_from_basket.gif" BORDER=0 ALT="Remove from Basket" ALIGN=LEFT>
                        </A>
                    </TD>
                </TR>
            </TABLE>
            </TD>
        </FORM>
        </TR>
    <% end if %>
</TABLE>
<% conn.close %>
<!--#INCLUDE FILE = "include/footer.asp" -->

</BODY>
</HTML>


