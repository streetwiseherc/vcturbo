<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   SHOPPER_DAY.ASP                                                         %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<!--#INCLUDE FILE="include/date_args2var.asp" -->
<% pageTitle = "New Shoppers by Month: " & imonth & "/" & iyear %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->


<FORM METHOD="POST" ACTION="shopper_day.asp">
    Time Period: <!--#INCLUDE FILE="include/picker_month.asp" --> 
                 <!--#INCLUDE FILE="include/picker_year.asp" -->
    <INPUT TYPE="SUBMIT" VALUE="Update">
</FORM>

<%
function ListRow() %>
    <TD VALIGN=TOP ALIGN=CENTER> <% = rsList("day") %>          </TD>
    <TD VALIGN=TOP ALIGN=CENTER> <% = rsList("num_shoppers") %> </TD>
<% end function

function ListHeader() %>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Day          </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> New Shoppers </TH>
<% end function
listElement = "shoppers during the month of"
ListSQL = "select { fn DAYOFMONTH(created) } day , count(shopper_id) num_shoppers from " & mscsTablePrefix & "shopper where  { fn YEAR(created) } = " & iyear & " and { fn MONTH(created) } = " & imonth & " group by { fn DAYOFMONTH(created) } order by 1"    
ListNoRows = "No shoppers found"
%>

<!--#INCLUDE FILE="include/list.asp" -->

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
