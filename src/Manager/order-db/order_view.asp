<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   ORDER_VIEW.ASP                                                          %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = ("Order '") & (mscsPage.HTMLEncode(CStr(Request("order_id")))) & ("'") %>


<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<%
order_id = Request("order_id")

sql = "select * from vcturbo_receipt where order_id = " & order_id
set rsReceipt = conn.Execute(sql)

sql = "select * from vcturbo_receipt_item where order_id = " & order_id & " order by item_id"
set rsReceiptItems = conn.Execute(sql)

if rsReceipt.EOF or rsReceiptItems.EOF then %>
    Order Not Found!
<% else %>
<BR>
<A HREF="<% = Replace(mscsShopperPath, ":1", rsReceipt("shopper_id").Value) %>">View Shopper </A></TD>
<br>
<br>
<TABLE>
    <TR>
        <% REM date:  %>
        <TD VALIGN=TOP>
            <B>Date:</B>
        </TD>
        <TD VALIGN=TOP>
            <% = FormatDateTime(CDate(rsReceipt("created").Value), vbShortDate) %>
        </TD>
    </TR>
    <TR>
        <% REM order status:  %>
        <TD VALIGN=TOP>
            <B>Status:</B>
        </TD>
        <TD VALIGN=TOP>
        </TD>
    </TR>
    <TR>
        <% REM Shopper ID %>
        <TD VALIGN="TOP">
            <B>Shopper ID:</B>
        </TD>
        <TD VALIGN="TOP">
            <% = mscsPage.HTMLEncode(rsReceipt("shopper_id").Value) %>
        </TD>
    </TR>
    <TR>
        <% REM bill to:  %>
        <TD VALIGN=TOP>
            <B>Bill To:</B>
        </TD>
        <TD VALIGN=TOP>
            <% = mscsPage.HTMLEncode(rsReceipt("bill_to_name").Value) %>
            <BR>
            <% = mscsPage.HTMLEncode(rsReceipt("bill_to_street").Value) %>
            <BR>
            <% = mscsPage.HTMLEncode(rsReceipt("bill_to_city").Value) %>, <% = mscsPage.Encode(rsReceipt("bill_to_state").Value) %> &nbsp <% = mscsPage.Encode(rsReceipt("bill_to_zip").Value) %>
        </TD>
        <% REM col spacer:  %>
        <TD WIDTH=16></TD>
        <% REM ship to:  %>
        <TD VALIGN=TOP>
            <B>Ship To:</B>
        </TD>
        <TD VALIGN=TOP>
            <% = mscsPage.HTMLEncode(rsReceipt("ship_to_name").Value) %>
            <BR>
            <% = mscsPage.HTMLEncode(rsReceipt("ship_to_street").Value) %>
            <BR>
            <% = mscsPage.HTMLEncode(rsReceipt("ship_to_city").Value) %>, <% = mscsPage.Encode(rsReceipt("ship_to_state").Value) %> &nbsp <% = mscsPage.Encode(rsReceipt("bill_to_zip").Value) %>
        </TD>
    </TR>
</TABLE>
<BR>
<TABLE CELLPADDING=3 CELLSPACING=0 BORDER=0>
    <% REM column headers:  %>
    <TR>
        <TH VALIGN=BOTTOM ALIGN=LEFT WIDTH=50>Sku</TH>
        <TH VALIGN=BOTTOM ALIGN=LEFT WIDTH=50>PFID</TH>
        <TH VALIGN=BOTTOM ALIGN=LEFT WIDTH=180>Item</TH>
        <TH VALIGN=BOTTOM ALIGN=CENTER WIDTH=30>Qty</TH>
        <TH VALIGN=BOTTOM ALIGN=CENTER WIDTH=60>List Price</TH>
        <TH VALIGN=BOTTOM ALIGN=CENTER WIDTH=60>Discount</TH>
        <TH VALIGN=BOTTOM ALIGN=CENTER WIDTH=60>Total</TH>
    <TR>
    <TR> <TD COLSPAN=9><HR></TD></TR>
    <% REM line items:  %>
	<% while Not(rsReceiptItems.EOF) %>
        <TR>
            <TD VALIGN=TOP ALIGN=LEFT><% = mscsPage.HTMLEncode(rsReceiptItems("sku").Value) %></TD>
            <TD VALIGN=TOP ALIGN=LEFT><% = mscsPage.HTMLEncode(rsReceiptItems("pfid").Value) %></TD>
            <TD VALIGN=TOP ALIGN=LEFT><% = mscsPage.HTMLEncode(rsReceiptItems("name").Value) %><% = mscsPage.HTMLEncode(rsReceiptItems("monogram").Value) %> <% = mscsPage.HTMLEncode(rsReceiptItems("attr_text").Value) %></TD> 
            <TD ALIGN=CENTER><% = mscsPage.HTMLEncode(rsReceiptItems("quantity").Value) %></TD>
            <TD ALIGN=CNETER><% = MSCSDataFunctions.Money(rsReceiptItems("product_list_price").Value) %></TD>
            <TD ALIGN=CENTER><% = MSCSDataFunctions.Money(rsReceiptItems("product_list_price").Value * rsReceiptItems("quantity").Value - rsReceiptItems("oadjust_adjustedprice").Value) %></TD>
            <TD ALIGN=CENTER><% = MSCSDataFunctions.Money(rsReceiptItems("oadjust_adjustedprice").Value) %></TD>
        </TR>
        <% rsReceiptItems.MoveNext %>
    <% wend %>
    <% REM divider:  %>
    <TR> <TD COLSPAN=8><HR></TD></TR>
    <% REM show subtotal:  %>
    <TR>
        <TD COLSPAN=4></TD>
        <TD COLSPAN=2 VALIGN=TOP ALIGN=LEFT>
            Subtotal:
        </TD>
        <TD ALIGN=RIGHT>
            <% = MSCSDataFunctions.Money(rsReceipt("oadjust_subtotal").Value) %>
        </TD>
    </TR>
    <% REM show shipping:  %>
    <TR>
        <TD COLSPAN=4></TD>
        <TD COLSPAN=2 VALIGN=TOP ALIGN=LEFT>
            Shipping:
        </TD>
        <TD ALIGN=RIGHT>
            <% = MSCSDataFunctions.Money(rsReceipt("shipping_total").Value) %>
        </TD>
    </TR>
    <% REM show tax:  %>
    <TR>
        <TD COLSPAN=4></TD>
        <TD COLSPAN=2 VALIGN=TOP ALIGN=LEFT>
            Tax:
        </TD>
        <TD ALIGN=RIGHT>
            <% = MSCSDataFunctions.Money(rsReceipt("tax_total").Value) %>
        </TD>
    </TR>
    <% REM divider:  %>
    <TR>
        <TD COLSPAN=3 HEIGHT=1></TD>
        <TD COLSPAN=4><HR></TD>
    </TR>
    <% REM show total:  %>
    <TR>
        <TD COLSPAN=4></TD>
        <TD COLSPAN=2 VALIGN=TOP ALIGN=LEFT>
            <B>Total:</B>
        </TD>
        <TD VALIGN=TOP ALIGN=RIGHT>
            <B><% = MSCSDataFunctions.Money(rsReceipt("total_total").Value) %></B>
        </TD>
    </TR>
</TABLE>

<% end if %>

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
