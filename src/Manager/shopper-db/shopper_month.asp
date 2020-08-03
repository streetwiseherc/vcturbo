<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   SHOPPER_MONTH.ASP                                                       %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<!--#INCLUDE FILE="include/date_args2var.asp" -->
<% pageTitle = "New Shoppers by Year: " & iyear %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->


<FORM METHOD="POST" ACTION="shopper_month.asp">
    Time Period: <!--#INCLUDE FILE="include/picker_year.asp" -->
    <INPUT TYPE="SUBMIT" VALUE="Update">
</FORM>

<%
function ListRow %>
    <TD VALIGN=TOP VALIGN=TOP ALIGN=CENTER> <% = rsList("month") %>        </TD>
    <TD VALIGN=TOP ALIGN=CENTER> <% = rsList("num_shoppers") %> </TD>
<% end function

function ListHeader() %>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Month        </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> New Shoppers </TH>
<%end function
ListNoRows = "No shoppers found"
ListSQL = "select { fn MONTHNAME(created) } month, count(shopper_id) num_shoppers from " & mscsTablePrefix & "shopper where { fn YEAR(created) } = " & iyear & " group by { fn MONTHNAME(created) } order by 1"
%>
<!--#INCLUDE FILE="include/list.asp" -->

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
