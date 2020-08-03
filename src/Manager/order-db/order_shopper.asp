<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   ORDER_SHOPPER.ASP                                                       %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<!--#INCLUDE FILE="include/date_args2var.asp" -->

<% pageTitle = "Orders by Shopper: " & imonth & "/" & iyear %>


<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<FORM METHOD="POST" ACTION="order_shopper.asp">
    Time Period: <!--#INCLUDE FILE="include/picker_month.asp" --> 
                 <!--#INCLUDE FILE="include/picker_year.asp" -->
    <INPUT TYPE="SUBMIT" VALUE="Update">
</FORM>

<%
function ListRow() %>
    <TD VALIGN=TOP ALIGN=CENTER> <% = RowCount %>     </TD>
    <TD VALIGN=TOP ALIGN=CENTER> <A HREF="<% = Replace(mscsShopperPath, ":1", mscsPage.URLEncode(rsList("shopper_id").Value)) %>"> <% = rsList("shopper_id") %> </A> </TD>
    <TD VALIGN=TOP ALIGN=CENTER> <A HREF="order_shopper_list.asp?shopper_id=<% = mscsPage.URLEncode(rsList("shopper_id").Value) %>"><% = rsList("num_orders") %> </A></TD>
    <TD VALIGN=TOP ALIGN=RIGHT > <% = MSCSDataFunctions.Money(rsList("total_sum").Value) %>  </TD>
    <TD VALIGN=TOP ALIGN=RIGHT > <% = MSCSDataFunctions.Money(rsList("total_avg").Value) %>  </TD>
    <TD VALIGN=TOP ALIGN=RIGHT > <% = MSCSDataFunctions.Money(rsList("total_max").Value) %>  </TD>
    <TD VALIGN=TOP ALIGN=RIGHT > <% = MSCSDataFunctions.Money(rsList("total_min").Value) %>  </TD>
<% end function
function ListHeader() %>
    <TH ALIGN=LEFT VALIGN=BOTTOM> #          </TH>
	<TH ALIGN=LEFT VALIGN=BOTTOM> Shopper Id </TH>
	<TH ALIGN=LEFT VALIGN=BOTTOM> Num Orders </TH>
	<TH ALIGN=LEFT VALIGN=BOTTOM> Total $    </TH>
	<TH ALIGN=LEFT VALIGN=BOTTOM> Avg $      </TH>
	<TH ALIGN=LEFT VALIGN=BOTTOM> Max $      </TH>
	<TH ALIGN=LEFT VALIGN=BOTTOM> Min $      </TH>
<% end function
listElemTemplate = "shopper_view.asp"
listElement = "shopper orders during the month of"
ListSQL = "select shopper_id, count(order_id) num_orders, sum(total_total) total_sum, avg(total_total) total_avg, max(total_total) total_max, min(total_total) total_min from vcturbo_receipt where { fn YEAR(created) } = " & iyear & "and { fn MONTH(created) } = " & imonth & " group by shopper_id order by 2"
%>
<!--#INCLUDE FILE="include/list.asp" -->

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
