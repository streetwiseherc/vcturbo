<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   ORDER_LIST.ASP                                                          %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "All Orders" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<%
function ListRow() %>
    <TD VALIGN=TOP ALIGN=CENTER> <% = RowCount %> </TD>
    <% order_id = rsList("order_id").Value %>
    <TD VALIGN=TOP ALIGN=CENTER> <A HREF="order_view.asp?order_id=<% = order_id %>"><% = mscsPage.HTMLEncode(order_id) %></A></TD>
    <TD VALIGN=TOP ALIGN=CENTER> <% = FormatDateTime(rsList("created").Value, vbShortDate) %> </TD>
    <TD VALIGN=TOP ALIGN=CENTER> <A HREF="<% = Replace(mscsShopperPath, ":1", mscsPage.URLEncode(rsList("shopper_id").Value)) %>"> <% = rsList("shopper_id").Value %></A></TD>
    <TD WIDTH=80  VALIGN=TOP ALIGN=RIGHT > <% = MSCSDataFunctions.Money(rsList("total_total").Value) %> </TD>
<% end function
function ListHeader() %>
    <TH ALIGN=LEFT VALIGN=BOTTOM> #          </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Order ID   </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Date       </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Shopper ID </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Total      </TH>
<% end function
listElemTemplate = "order_view.asp"
listSQL = "SELECT * FROM vcturbo_receipt ORDER BY order_id"
%>
<!--#INCLUDE FILE="include/list.asp" -->

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
