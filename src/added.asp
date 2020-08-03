<%@ LANGUAGE = vbscript %>
<!--#INCLUDE FILE = "include/shop.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   UPSELL.ASP                                                             #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>

<HTML>

<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>" BACKGROUND="<% = background %>">

<% hd_image = "heading_products.gif "%>
<% alt = "Confirm Basket Page" %>
<!--#INCLUDE FILE = "include/header.asp" -->

<%
REM Get the item just added
Dim rsItem
Dim pfid 
Dim monogram
Dim quantity
Dim sku
Dim item_id
Dim pfReplace

set rsItem = Conn.Execute("select item_id, pfid, name, sku, attr_text, monogram, quantity from vcturbo_basket_item where shopper_id = '" & mscsShopperId & "' order by item_id desc")
pfid       = rsItem("pfid").Value
monogram   = rsItem("monogram").Value
quantity   = rsItem("quantity").Value
sku        = rsItem("sku").Value
item_id    = rsItem("item_id").Value

pfReplace  = Replace(pfid, "'", "''")
%>

<H2>Thank you for interest in <% = rsItem("name").Value %>.<BR> It has been added to your basket.</H2>
<% Dim rsUpSell %>
<% set rsUpSell = Conn.Execute("select related_pfid, name from vcturbo_promo_upsell, vcturbo_product_family where vcturbo_promo_upsell.pfid = '" & pfReplace & "' and vcturbo_product_family.pfid = vcturbo_promo_upsell.related_pfid") %>

<% if Not rsUpSell.EOF then %>
     <br>
     Perhaps you would be interested in the following product(s) instead?
     <UL>
     <% while Not rsUpSell.EOF %>               
           <li><A HREF="<% = baseURL("product_alt.asp") & mscsPage.URLShopperArgs("pfid", rsUpSell("related_pfid").Value, "monogram", monogram, "quantity", quantity, "item_id", item_id) %>"><% = rsUpSell("name").Value %></A>
           <% rsUpSell.MoveNext %>
     <% wend %>
     </UL>
<% end if %>
<% Dim rsCrossSell %>                           
<% set rsCrossSell = Conn.Execute("select related_pfid, name from vcturbo_promo_cross, vcturbo_product_family where vcturbo_promo_cross.pfid = '" & pfReplace & "' and vcturbo_product_family.pfid = vcturbo_promo_cross.related_pfid") %>

<% if Not rsCrossSell.EOF then %>
     <br>
     Are you interested in information about the following related product(s)?
     <UL>
         <% while Not rsCrossSell.EOF %>
              <li><A HREF="<% = baseURL("product.asp") & mscsPage.URLShopperArgs("pfid", rsCrossSell("related_pfid").Value) %>"><% = rsCrossSell("name").Value %></A>
              <% rsCrossSell.MoveNext %>
         <% wend %>
     </UL>
<% end if %>

<br>

<A HREF = "<% = pageURL("basket.asp") %>"><IMG SRC="assets/images/button_basket_up.gif" BORDER=0 ALT="View Basket"></A>
<A HREF = "<% = pageURL("listing.asp") %>"><IMG SRC="assets/images/button_catalog_up.gif" BORDER=0 ALT="Continue"></A>

<br>
<% conn.close %>
<!--#INCLUDE FILE = "include/footer.asp" -->

</BODY>
</HTML>

