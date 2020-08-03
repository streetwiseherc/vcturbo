<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   ORDER_PRODUCT.ASP                                                       %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<!--#INCLUDE FILE="include/date_args2var.asp" -->
<% pageTitle = "Orders by Product: " & imonth & "/" & iyear %>


<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<FORM METHOD="POST" ACTION="order_product.asp">
    Time Period: <!--#INCLUDE FILE="include/picker_month.asp" --> 
                 <!--#INCLUDE FILE="include/picker_year.asp" -->
    <INPUT TYPE="SUBMIT" VALUE="Update">
</FORM>

<%
function ListRow() %>
        <TD VALIGN=TOP ALIGN=CENTER> <% = RowCount %>    </TD>
        <TD VALIGN=TOP ALIGN=CENTER> <A HREF="<% = Replace(mscsProductPath, ":1", mscsPage.URLEncode(rsList("pfid").Value)) %>"> <% = rsList("pfid") %> </A> </TD>
        <TD VALIGN=TOP ALIGN=LEFT  > <% = rsList("name") %>   </TD>
        <TD VALIGN=TOP ALIGN=CENTER> <% = rsList("num_items") %> </TD>
        <TD VALIGN=TOP ALIGN=RIGHT > <% = MSCSDataFunctions.Money(rsList("price_sum")) %> </TD>
        <TD VALIGN=TOP ALIGN=RIGHT > <% = MSCSDataFunctions.Money(rsList("price_avg")) %> </TD>
        <TD VALIGN=TOP ALIGN=RIGHT > <% = MSCSDataFunctions.Money(rsList("price_max")) %> </TD>
        <TD VALIGN=TOP ALIGN=RIGHT > <% = MSCSDataFunctions.Money(rsList("price_min")) %> </TD>
        <TD VALIGN=TOP ALIGN=CENTER> <% = rsList("qty_sum") %>   </TD>
        <TD VALIGN=TOP ALIGN=CENTER> <% = rsList("qty_avg") %>   </TD>
        <TD VALIGN=TOP ALIGN=CENTER> <% = rsList("qty_max") %>   </TD>
        <TD VALIGN=TOP ALIGN=CENTER> <% = rsList("qty_min") %>   </TD>

<% end function
function ListHeader() %>
		<TH ALIGN=LEFT VALIGN=BOTTOM> #               </TH>
		<TH ALIGN=LEFT VALIGN=BOTTOM> PF Id           </TH>
		<TH ALIGN=LEFT VALIGN=BOTTOM> Name            </TH>
		<TH ALIGN=LEFT VALIGN=BOTTOM> Num Orders      </TH>
		<TH ALIGN=LEFT VALIGN=BOTTOM> Total $         </TH>
		<TH ALIGN=LEFT VALIGN=BOTTOM> Avg $           </TH>
		<TH ALIGN=LEFT VALIGN=BOTTOM> Max $           </TH>
		<TH ALIGN=LEFT VALIGN=BOTTOM> Min $           </TH>
		<TH ALIGN=LEFT VALIGN=BOTTOM> Items/Order     </TH>
		<TH ALIGN=LEFT VALIGN=BOTTOM> Avg Items/Order </TH>
		<TH ALIGN=LEFT VALIGN=BOTTOM> Max Items/Order </TH>
		<TH ALIGN=LEFT VALIGN=BOTTOM> Min Items/Order </TH>
<% end function
listElemTemplate = "product_edit.asp"
sql = "select item.pfid, item.name, count(item.pfid) num_items, sum(item.oadjust_adjustedprice) price_sum, avg(item.oadjust_adjustedprice) price_avg, max(item.oadjust_adjustedprice) price_max, min(item.oadjust_adjustedprice) price_min, sum(item.quantity) qty_sum, avg(item.quantity) qty_avg, max(item.quantity) qty_max, min(item.quantity) qty_min "
sql = sql + "from vcturbo_receipt receipt, vcturbo_receipt_item item "
sql = sql + "where receipt.order_id = item.order_id and { fn YEAR(receipt.created) } = " & iyear & "and { fn MONTH(receipt.created) } = " & imonth & " group by item.pfid, item.name order by 3"
ListSQL = sql
%>
<!--#INCLUDE FILE="include/list.asp" -->

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
