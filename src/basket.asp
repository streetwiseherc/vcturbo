<%@ LANGUAGE=vbscript enablesessionstate=false %>
<!--#INCLUDE FILE = "include/shop.asp" -->
<!--#INCLUDE FILE = "include/util.asp" -->
<!--#INCLUDE FILE = "include/order.asp" -->
<% Response.ExpiresAbsolute=#Jan 01, 1980 00:00:00# %>
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   BASKET.ASP                                                             #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>

<%
REM Run the basic plan
Dim nItemCount
Dim mscsOrderForm 

Set mscsOrderForm = UtilRunBasket(mscsShopperId)
nItemCount = mscsOrderForm.Items.Count
%>

<HTML>

<HEAD>
    <TITLE>Shopping Basket</TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>
<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>"BACKGROUND="<% = background %>">
<OBJECT RUNAT=Server ID=errorList PROGID="Commerce.SimpleList"></OBJECT>
<% hd_image = "heading_basket.gif" %>
<% alt = "Shopping Basket" %>
<!--#INCLUDE FILE = "include/header.asp" -->


<% if mscsOrderForm.[_Basket_Errors].Count > 0 then  %>
    <UL>
        <% for each errorStr in mscsOrderForm.[_Basket_Errors] %>
            <li><% = errorStr %></li>
        <% next %>
    </UL>
    <P>
<% end if %>

<% if nItemCount = 0 then %>
    <BLOCKQUOTE><STRONG>Your basket is empty.</STRONG></BLOCKQUOTE>
    <P>
<% else %>
    You have <% = nItemCount %> item<% if nItemCount > 1 then %>s<% end if %> in your basket.
    <P>
    To change or remove an item in the basket, click on its description.
    <P>
    <TABLE WIDTH=600>
        <TR><TD colspan=3>
            <TABLE>
                <TR>
                    <TH ALIGN=LEFT>Item</TH>
                    <TH>&nbsp;&nbsp;</TH>
                    <TH ALIGN=LEFT>Quantity</TH>
                    <TH>&nbsp;&nbsp;</TH>
                    <TH ALIGN=LEFT>Description</TH>
                    <TH>&nbsp;&nbsp;</TH>
                    <TH ALIGN=CENTER>Unit Price</TH>
                    <TH>&nbsp;&nbsp;</TH>
                    <TH ALIGN=CENTER>Promo Disc.</TH>
                    <TH>&nbsp;&nbsp;</TH>
                    <TH ALIGN=CENTER>Total Price</TH>
                </TR>
                <TR><TD colspan=11><HR></TD>
                </TR>
                <% Dim rowidx %>
                <% Dim row_orderitem %>
                <% Dim monogram %>
                <% Dim attr_text %>
                <% rowidx = 1 %>
                <% for each row_orderitem in mscsOrderForm.Items %>
                            <TR>
                            <TD><% = mscsPage.HTMLEncode(row_orderitem.sku) %></TD> 
                            <TD>&nbsp;&nbsp;</TD>
                            <TD><% = mscsPage.HTMLEncode(row_orderitem.quantity) %></TD> 
                            <TD>&nbsp;&nbsp;</TD>
                            <TD>
                                <A HREF = "<% = baseURL("product_edit.asp") &mscsPage.URLShopperArgs("pfid", CStr(row_orderitem.pfid), "sku", cstr(row_orderitem.sku) , "item_id" , row_orderitem.item_id , "quantity" ,  cstr(row_orderitem.quantity), "monogram", monogram) %>"> <% = row_orderitem.name %></A>
                            </TD>
                            <TD>&nbsp;&nbsp;</TD> 
                            <TD ALIGN=RIGHT><% = MSCSDataFunctions.Money(row_orderitem.[_iadjust_regularprice]) %></TD> 
                            <TD>&nbsp;&nbsp;</TD> 
                            <TD ALIGN=RIGHT><% = MSCSDataFunctions.Money(row_orderitem.[_oadjust_discount]) %></TD> 
                            <TD>&nbsp;&nbsp;</TD>
                            <TD ALIGN=RIGHT><% = MSCSDataFunctions.Money(row_orderitem.[_oadjust_adjustedprice]) %></TD>
                            <% attr_text = OrderBasketAttrText(row_orderitem) %>
                            <% if attr_text <> "" then %>
                                <TR>
                                    <TD colspan=4>&nbsp;&nbsp;</TD>
                                    <TD colspan=5>
                                        <% = mscsPage.HTMLEncode(attr_text) %>
                                    </TD>
                                </TR>
                            <% end if %>
                    <% rowidx = rowidx + 1 %>
                <% next %>
                <TR><TD colspan=11><hr></TD>
                </TR>
                <TR>
                    <TD colspan=8>&nbsp;</TD>
                    <TD ALIGN=RIGHT><STRONG>Subtotal:</STRONG></TD>
                    <TD>&nbsp;&nbsp;</TD>
                    <TD ALIGN=RIGHT><% = MSCSDataFunctions.Money(mscsOrderForm.[_oadjust_subtotal]) %></TD>
                </TR>
            </TABLE>
            </TD>
        </TR>
        <TR><TD colspan=5>&nbsp;</TD></TR>
    </TABLE>

    <TABLE WIDTH=600>
        <TR>
            <TD>
                <FORM METHOD=POST ACTION="<% = pageURL("xt_orderform_clearitems.asp") %>">
                    <INPUT TYPE="Image" VALUE="Clear Order" SRC="assets/images/btn_empty_basket.gif" BORDER=0 ALT="Clear Basket">
                </FORM>
            </TD>
            <TD>
                <FORM METHOD=POST ACTION = "<% = pageURL("listing.asp") %>">
                    <INPUT TYPE="Image" VALUE="Continue Shopping" SRC="assets/images/btn_continue_shopping.gif" WIDTH=84 HEIGHT=51 BORDER=0 ALT="Continue Shopping">
                </FORM>
            </TD>
            <TD>
                <FORM METHOD=POST ACTION = "<% = pageURL("checkout-ship.asp") %>">
                    <INPUT TYPE="Image" VALUE="Purchase" SRC="assets/images/btn_purchase_now.gif" BORDER=0 ALT="Purchase Now">
                </FORM>
            </TD>
        </TR>
    </TABLE>
       
<% end if %>
<% conn.close %>
<!--#INCLUDE FILE = "include/footer.asp" -->

</BODY>
</HTML>
