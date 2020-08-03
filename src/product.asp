<%@ LANGUAGE = vbscript %>
<!--#INCLUDE FILE = "include/shop.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   PRODUCT.ASP                                                            #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>

<HTML>

<%
Dim pfid
Dim rsProduct
Dim name

pfid = Request("pfid")   
set rsProduct = Conn.Execute("select pfid, monogramable, list_price, name, image_filename, short_description, long_description from vcturbo_product_family where pfid = '" & Replace(pfid, "'", "''") & "'")

if rsProduct.EOF then
    name = "Unknown"
else
    name = rsProduct("name").Value
end if
%>

<HEAD>
     <TITLE>Product <% = mscsPage.HTMLEncode(name) %></TITLE>
     <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>
<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>" BACKGROUND="<% = background %>">

<%  hd_image = "heading_products.gif "
    alt      = "Product Page" %>
<!--#INCLUDE FILE = "include/header.asp" -->


<TABLE WIDTH=600>
    <% if rsProduct.EOF then %>
        <TR>
            <TD>Product not available</TD>
        </TR>
    <% else %>
        <TR>
            <TD colspan=3><H2><% = mscsPage.HTMLEncode(name) %></H2></TD>
        </TR>
        <TR>
            <TD WIDTH=30% VALIGN="top">
                <% if not isnull(rsProduct("image_filename").Value) then %>
                    <IMG SRC="assets/prodimg/<% = rsProduct("image_filename").Value %>">
                <% else %>
                    Image not available
                <% end if %>
            </TD>
            <TD VALIGN=TOP WIDTH=400 colspan=2><% = mscsPage.HTMLEncode(rsProduct("short_description").Value) %></TD>
        </TR>
        <TR>
            <TD VALIGN=TOP WIDTH=400 COLSPAN=3><% = mscsPage.HTMLEncode(rsProduct("long_description").Value) %></TD>
        </TR>
        <TR>
            <FORM METHOD=POST ACTION="<% = pageURL("xt_orderform_additem.asp") %>">
                <INPUT TYPE = HIDDEN NAME = "pfid" VALUE = <% = rsProduct("pfid").Value %>>
                <TD valign=top><STRONG>Our Price:</STRONG>&nbsp;<FONT SIZE="+2"><% = MSCSDataFunctions.Money(CLng(rsProduct("list_price").Value)) %></FONT></TD>
                <TD VALIGN=TOP ALIGN=RIGHT><!-- #INCLUDE FILE = "include/prd_body.asp" --></TD>
                <TD VALIGN=TOP ALIGN=LEFT>
                    <INPUT TYPE=image SRC="assets/images/btn_add_2_basket_up.gif" BORDER=0 ALT="Add to Basket" ALIGN=left>
                </TD>
            </FORM>                                                  
        </TR>
    <% end if %>
</TABLE>
<% conn.close %>
<!--#INCLUDE FILE = "include/footer.asp" -->

</BODY>
</HTML>
