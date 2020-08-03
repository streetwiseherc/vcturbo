<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   SHOPPER_VIEW.ASP                                                        %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "Shopper '" &  mscsPage.HTMLEncode(Request("shopper_id")) & "'" %>


<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<%
shopper_id = Request("shopper_id")
%>
<%
sql = "select name, street, city, state, zip, country, email, phone from vcturbo_shopper where shopper_id = '" & shopper_id & "'"
set rsShopper = conn.Execute(sql)
if Not(rsShopper.EOF) then %>
<BR>
<A HREF="<% = Replace(mscsReceiptsPath, ":1", Request("shopper_id")) %>">View Receipts</A><br>
<A HREF="<% = Replace(mscsBasketPath, ":1", Request("shopper_id")) %>">View Basket</A><br>
<BR>
<TABLE>
    <TR>
        <% REM basic:  %>
        <TD VALIGN=TOP>
            <B><% = mscsPage.HTMLEncode(rsShopper("name").Value) %></B>
            <BR>
            <% = mscsPage.HTMLEncode(rsShopper("street").Value) %>
            <BR>
            <% = mscsPage.HTMLEncode(rsShopper("city").Value) %>, <% = mscsPage.HTMLEncode(rsShopper("state").Value) %> &nbsp <% = mscsPage.HTMLEncode(rsShopper("zip").Value) %>
            <BR>
            <% = mscsPage.HTMLEncode(rsShopper("country").Value) %>
            <BR>
            <% = mscsPage.HTMLEncode(rsShopper("phone").Value) %>
            <BR>
            <% = mscsPage.HTMLEncode(rsShopper("email").Value) %>
            
        </TD>
    </TR>
</TABLE>
<BR>
<FORM METHOD=POST ACTION="confirm.asp">
    <INPUT TYPE="HIDDEN" NAME="confirm_action" VALUE="xt_data_delete.asp">
    <INPUT TYPE="HIDDEN" NAME="confirm_message" VALUE="Delete the shopper '<% = mscsPage.HTMLEncode(shopper_id) %>'?">
    <INPUT TYPE="HIDDEN" NAME="shopper_id" VALUE="<% = mscsPage.HTMLEncode(shopper_id) %>">
    <INPUT TYPE="SUBMIT" VALUE="Delete Shopper...">
</FORM>
<BR>
<% else %>
    <P>
    <STRONG>Shopper "<% = Request("shopper_id") %>" not found</STRONG>
<% end if %>

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
