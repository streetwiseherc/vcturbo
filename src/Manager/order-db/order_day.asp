<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   ORDER_DAY.ASP                                                           %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<!--#INCLUDE FILE="include/date_args2var.asp" -->

<% pageTitle = "New Orders by Month: " & imonth & "/" & iyear %>


<!--#INCLUDE FILE="include/mgmt_header.asp" -->


<% REM   body: %>
<FORM METHOD="POST" ACTION="order_day.asp">
    Time Period: <!--#INCLUDE FILE="include/picker_month.asp" --> 
                 <!--#INCLUDE FILE="include/picker_year.asp" -->
    <INPUT TYPE="SUBMIT" VALUE="Update">
</FORM>

<%
function ListRow() %>
    <TD VALIGN=TOP ALIGN=CENTER> <% = rsList("day") %>        </TD>
    <TD VALIGN=TOP ALIGN=CENTER> <% = rsList("num_orders") %> </TD>
    <TD VALIGN=TOP ALIGN=RIGHT > <% = MSCSDataFunctions.Money(Cstr(rsList("total_sum"))) %>  </TD>
    <TD VALIGN=TOP ALIGN=RIGHT > <% = MSCSDataFunctions.Money(Cstr(rsList("total_avg"))) %>  </TD>
    <TD VALIGN=TOP ALIGN=RIGHT > <% = MSCSDataFunctions.Money(Cstr(rsList("total_max"))) %>  </TD>
    <TD VALIGN=TOP ALIGN=RIGHT > <% = MSCSDataFunctions.Money(Cstr(rsList("total_min"))) %>  </TD>

<% end function
function ListHeader() %>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Day        </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Num Orders </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Total $    </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Avg $      </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Max $      </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Min $      </TH>
<% end function
ListSQL = "select { fn DAYOFMONTH(created) } day, count(order_id) num_orders, sum(total_total) total_sum, avg(total_total) total_avg, max(total_total) total_max, min(total_total) total_min from vcturbo_receipt where { fn YEAR(created) } = " & iyear & " and { fn MONTH(created) } = " & imonth & " group by { fn DAYOFMONTH(created) } order by 1"
%>
<!--#INCLUDE FILE="include/list.asp" -->

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
