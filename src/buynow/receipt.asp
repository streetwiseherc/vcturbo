<%@ LANGUAGE = VBScript %>

<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   RECEIPT.ASP                                                             %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<!--#INCLUDE FILE="include/shop.asp" -->
<!--#INCLUDE FILE="include/util.asp" -->
<!--#INCLUDE FILE="include/order.asp" -->

<HTML>

<HEAD>
    <TITLE><% = displayName %> : Receipt</TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>

<BODY LEFTMARGIN=16 TOPMARGIN=16 LANGUAGE="Javascript" BACKGROUND="../assets/images/bnbackground.gif">

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

<H2>
    Buy now Receipt
    <FONT color="#000000" size="4">
        <B>
            <BR>
            <P>
            <P>
            Receipt #
            <% = Request("order_id") %>
        </B>
    </FONT>
</h2>

<TABLE WIDTH=100% BORDER="0" CELLPADDING="0" CELLSPACING="0">
<% if nItems = 0 then %>
    <B>Order Not Found</B>
<% else %>
        <TR>
            <TD VALIGN="top">
                <P>
                <STRONG>
                    Paying With:
                </STRONG>
                <BR>
                <% = mscsPage.HTMLEncode(receipt.[_cc_name]) %>'s&nbsp
                <% = mscsPage.HTMLEncode(receipt.[_cc_type]) %>
                <P>
                <STRONG>
                    Shipping To:
                </STRONG>
                <BR>
                <% = mscsPage.HTMLEncode(receipt.ship_to_name) %>  <BR>
                <% = mscsPage.HTMLEncode(receipt.ship_to_street) %> <BR>
                <% = mscsPage.HTMLEncode(receipt.ship_to_city) %>, 
                <% = mscsPage.HTMLEncode(receipt.ship_to_state) & " " %>
                <% = mscsPage.HTMLEncode(receipt.ship_to_zip) %>
                <P>
            </TD>
      
            <TD VALIGN="top">
                <STRONG>
                    Description:
                </STRONG>
                <BR>
                <FONT size="2">
                    <% REM line items:  %>
                    <% Dim row_item %>
                    <% for each row_item in items %>
                             
                    <% = mscsPage.HTMLEncode(row_item.quantity) %>&nbsp
                    <% = mscsPage.HTMLENcode(row_item.name) %>(s)
                    <% = mscsPage.HTMLEncode(row_item.attr_text) %>
                    
                    <% next %>
            </FONT>
            <BR>
                        <BR>
                            <TABLE BORDER="1" CELLPADDING="2" CELLSPACING="1">
                                <TR>
                                    <TD>
                                        Subtotal
                                    </TD>
                                    <TD ALIGN="right" WIDTH="75">
                                        <% = MSCSDataFunctions.Money(CLng(receipt.[_oadjust_subtotal])) %>
                                    </TD>
                                </TR>
                                <TR>
                                    <TD>
                                        Shipping
                                    </TD>
                                    <TD ALIGN="right">
                                        <% = MSCSDataFunctions.Money(CLng(receipt.[_shipping_total])) %>
                                    </TD>
                                </TR>
                                <TR>
                                    <TD>
                                        Tax
                                    </TD>
                                    <TD ALIGN="right">
                                        <% = MSCSDataFunctions.Money(CLng(receipt.[_tax_total])) %>
                                    </TD>
                                </TR>
                                <TR>
                                    <TD>
                                        Total
                                    </TD>
                                    <TD ALIGN="right">
                                        <% = MSCSDataFunctions.Money(CLng(receipt.[_total_total])) %>
                                    </TD>
                                </TR>
                            </TABLE>
                        </TD>
                    </TR>
<% end if %>
    <TR>      
        <TD colspan="3" ALIGN="middle">&nbsp</TD>
    </TR>
    <TR>
        <TD ALIGN="left"><H3>Step 4 of 4</H3></TD>
        <TD colspan="2" ALIGN="middle" VALIGN="bottom">
            <FORM>
                <% REM A print command would go here.... %>
                <INPUT TYPE="BUTTON" VALUE="Finished" ONCLICK="window.close()" >
            </FORM>
        </TD>
    </TR>
</TABLE>
</BODY>

</HTML>
