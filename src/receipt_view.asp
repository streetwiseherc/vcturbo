<!--#INCLUDE FILE="include/shop.asp" -->
<!--#INCLUDE FILE="include/order.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   RECEIPT_VIEW.ASP                                                          %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<HTML>
<HEAD>
<TITLE>Purchase Receipt</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>

<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>" BACKGROUND="<% = background %>">

<% hd_image = "heading_checkout.gif" %>
<% alt = "Checkout - Receipt"%>
<!--#INCLUDE FILE = "include/header.asp" -->

<% REM -- page vars:  %>
<% Dim pageTitle %>
<% pageTitle = ("Order '") & (CStr(Request("order_id"))) & ("'") %>

<%
Dim receipt 
set receipt = OrderGetReceipt(mscsShopperId, Request("order_id"))

Dim nitems
if IsEmpty(receipt) then
    nitems = 0
else
    Dim items
    set items = receipt.items
    nitems = items.Count
end if
%>

<BR>

<% if nItems = 0 then %>
    <B>Order Not Found</B>
<% else %>
    <TABLE WIDTH="600">
            <TR>
                <TD VALIGN=TOP>
                    <B>Order Id: </B>
                </TD>
                <TD VALIGN=TOP>
                    <% = Request("order_id") %>
                </TD>
            </TR>
            <TR>
                <% REM date:  %>
                <TD VALIGN=TOP>
                    <B>Date:</B>
                </TD>
                <TD VALIGN=TOP>
                    <% = FormatDateTime(CDate(receipt.created),vbShortDate) %>
                </TD>
            </TR>
            <TR>
                <% REM bill to:  %>
                <TD VALIGN=TOP>
                    <B>Bill To:</B>
                </TD>
                <TD VALIGN=TOP>
                    <% = mscsPage.HTMLEncode(receipt.bill_to_name) %>
                    <BR>
                    <% = mscsPage.HTMLEncode(receipt.bill_to_street) %>
                    <BR>
                    <% = mscsPage.HTMLEncode(receipt.bill_to_city) %>, <% = mscsPage.HTMLEncode(receipt.bill_to_state) %> &nbsp <% = mscsPage.HTMLEncode(receipt.bill_to_zip) %>
                </TD>
                
                <% REM col spacer:  %>
                <TD WIDTH=16></TD>

                <% REM ship to:  %>
                <TD VALIGN=TOP>
                    <B>Ship To:</B>
                </TD>
                <TD VALIGN=TOP>
                    <% = mscsPage.HTMLEncode(receipt.ship_to_name) %>
                    <BR>
                    <% = mscsPage.HTMLEncode(receipt.ship_to_street) %>
                    <BR>
                    <% = mscsPage.HTMLEncode(receipt.ship_to_city) %>, <% = mscsPage.HTMLEncode(receipt.ship_to_state) %> &nbsp <% = mscsPage.HTMLEncode(receipt.bill_to_zip) %>
                </TD>
            </TR>
    </TABLE>

    <BR>

    <TABLE WIDTH="600" CELLPADDING="0" CELLSPACING="3" BORDER="0">
            <% REM column headers:  %>
            <TR>
                <TH VALIGN=BOTTOM ALIGN=LEFT   WIDTH=50> Sku       </TH>
                <TH VALIGN=BOTTOM ALIGN=LEFT   WIDTH=180>Item      </TH>
                <TH VALIGN=BOTTOM ALIGN=CENTER WIDTH=30> Qty       </TH>
                <TH VALIGN=BOTTOM ALIGN=CENTER WIDTH=60> List Price</TH>
                <TH VALIGN=BOTTOM ALIGN=CENTER WIDTH=60> Disc      </TH>
                <TH VALIGN=BOTTOM ALIGN=CENTER WIDTH=60> Total     </TH>
                </FONT>
            </TR>
            <TR>
                <TD COLSPAN=6 BGCOLOR=black HEIGHT="2"><IMG SRC="manager/MSCS_Images/extras/dot_clear.gif" WIDTH=1 HEIGHT=1></TD>
            </TR>
            
            <% REM line items:  %>
            <% Dim row_item %>
            <% for each row_item in items %>
                
            <TR>
                <TD ALIGN=LEFT> <% = mscsPage.HTMLEncode(row_item.sku) %></TD>
                <TD ALIGN=LEFT><B><% = mscsPage.HTMLEncode(row_item.name) %></B><% = " " + mscsPage.HTMLEncode(OrderBasketAttrText(row_item)) %></TD>
                <TD ALIGN=CENTER> <% = mscsPage.HTMLEncode(row_item.quantity) %></TD>
                <TD ALIGN=CENTER> <% = MSCSDataFunctions.Money(CLng(row_item.[_iadjust_regularprice])) %></TD>
                <TD ALIGN=CENTER> <% = MSCSDataFunctions.Money(row_item.[_iadjust_currentprice] * row_item.quantity - row_item.[_oadjust_adjustedprice]) %></TD>
                <TD ALIGN=CENTER> <% = MSCSDataFunctions.Money(CLng(row_item.[_oadjust_adjustedprice])) %></TD>
            </TR>
            <% next %>
      
            <% REM divider:  %>
            <TR>
                <TD COLSPAN=6 BGCOLOR=black HEIGHT="2"><IMG SRC="manager/MSCS_Images/extras/dot_clear.gif" WIDTH=1 HEIGHT=1></TD>
            </TR>
            <% REM show subtotal:  %>
            <TR>
                <TD COLSPAN=3></TD>
                <TD COLSPAN=2 VALIGN=TOP ALIGN=LEFT>
                    Subtotal:
                </TD>
                <TD ALIGN=CENTER>
                    <% = MSCSDataFunctions.Money(CLng(receipt.[_oadjust_subtotal])) %>
                </TD>
            </TR>

            <% REM show shipping:  %>
            <TR>
                <TD COLSPAN=3></TD>
                <TD COLSPAN=2 VALIGN=TOP ALIGN=LEFT>
                    Shipping:
                </TD>
                <TD ALIGN=CENTER>
                    <% = MSCSDataFunctions.Money(CLng(receipt.[_shipping_total])) %>
                </TD>
            </TR>

            <% REM show tax:  %>
            <TR>
                <TD COLSPAN=3></TD>
                <TD COLSPAN=2 VALIGN=TOP ALIGN=LEFT>
                    Tax:
                </TD>
                <TD ALIGN=CENTER>
                    <% = MSCSDataFunctions.Money(CLng(receipt.[_tax_total])) %>
                </TD>
            </TR>

            <% REM divider:  %>
            <TR>
                <TD COLSPAN=3 HEIGHT=1></TD>
                <TD COLSPAN=4 BGCOLOR=black HEIGHT="2"><IMG SRC="manager/MSCS_Images/extras/dot_clear.gif" WIDTH=1 HEIGHT=1></TD>
            </TR>

            <% REM show total:  %>
            <TR>
                <TD COLSPAN=3></TD>
                <TD COLSPAN=2 VALIGN=TOP ALIGN=LEFT>
                    <B>Total:</B>
                </TD>
                <TD VALIGN=TOP ALIGN=CENTER>
                    <B><% = MSCSDataFunctions.Money(CLng(receipt.[_total_total])) %></B>
                </TD>
            </TR>

    </TABLE>
<% end if %>    

<!--#INCLUDE FILE="include/footer.asp" -->


</BODY>
</HTML>

