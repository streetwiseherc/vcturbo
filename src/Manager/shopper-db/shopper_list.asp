<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   SHOPPER_LIST.ASP                                                        %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "All Shoppers" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<P>

<%
function ListRow %>
    <TD VALIGN=TOP ALIGN=CENTER> <% = RowCount %> </TD>
    <TD VALIGN=TOP ALIGN=LEFT  > <A HREF="shopper_view.asp?<% = mscsPage.URLArgs("shopper_id", Cstr(rsList("shopper_id"))) %>"> <% = rsList("shopper_id")%> </A> </TD>
    <TD VALIGN=TOP ALIGN=LEFT  > <% = rsList("name") %>   </TD>
<% end function

function ListHeader() %>
    <TH ALIGN=LEFT VALIGN=BOTTOM> #          </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Shopper ID </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Name       </TH>
<% end function

ListNoRows = "No shoppers found"
ListSQL = "SELECT * FROM vcturbo_shopper ORDER BY shopper_id"
%>
<!--#INCLUDE FILE="include/list.asp" -->

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
