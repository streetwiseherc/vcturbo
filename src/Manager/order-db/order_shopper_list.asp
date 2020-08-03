<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   ORDER_SHOPPER_LIST.ASP                                                  %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% shopper_id = Request("shopper_id") %>
<% pageTitle = "Orders for shopper: " & shopper_id %>


<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<BR>
<A HREF="<% = Replace(mscsShopperpath, ":1", shopper_id) %>">View Shopper</A><br>
<BR>

<%
function ListRow() %>
    <TD VALIGN=TOP ALIGN=CENTER> <% = RowCount %> </TD>
    <% order_id = rsList("order_id").Value %>
    <TD VALIGN=TOP ALIGN=CENTER> <A HREF="order_view.asp?order_id=<% = order_id %>"><% = mscsPage.HTMLEncode(order_id) %></A></TD>
    <TD VALIGN=TOP ALIGN=CENTER> <% = FormatDateTime(rsList("created").Value, vbShortDate) %> </TD>
    <TD WIDTH=80  VALIGN=TOP ALIGN=RIGHT > <% = MSCSDataFunctions.Money(rsList("total_total").Value) %> </TD>
<% end function
function ListHeader() %>
    <TH ALIGN=LEFT VALIGN=BOTTOM> #          </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Order ID   </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Date       </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Total      </TH>
<% end function
listElemTemplate = "order_view.asp"
ListSQL = "SELECT * FROM vcturbo_receipt where shopper_id = '" & shopper_id & "' ORDER BY order_id"
%>
<!--#INCLUDE FILE="include/list.asp" -->

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
