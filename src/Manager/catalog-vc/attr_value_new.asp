<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   ATTR_VALUE_NEW.ASP                                                      %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "New Attribute Value" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<TABLE>
    <FORM METHOD="POST"
        ACTION="xt_data_add_update.asp">
        <INPUT TYPE="HIDDEN" NAME="type"            VALUE="attribute value">
        <INPUT TYPE="HIDDEN" NAME="goto"            VALUE="attr_value_list.asp">
        <INPUT TYPE="HIDDEN" NAME="attribute_id"    VALUE="<% = Request("attribute_id") %>">
        <INPUT TYPE="HIDDEN" NAME="attribute_index" VALUE="<% = Request("attribute_index") %>">

    <TR>
        <TD VALIGN=TOP>PFId:</TD>
        <TD VALIGN=TOP><STRONG><% = mscsPage.HTMLEncode(Request("pfid")) %></STRONG></TD>
    </TR>
        <INPUT TYPE="HIDDEN" NAME="pfid" VALUE="<% = Request("pfid") %>">
    <TR>
        <TD VALIGN=TOP>Attribute Name:</TD>
        <TD VALIGN=TOP><STRONG><% = Request("attr_name") %></STRONG></TD>
    </TR>
    <TR>
        <% REM label:  %>
        <TD VALIGN=TOP>
            * Name:
        </TD>
        <% REM value:  %>
        <TD VALIGN=TOP>
            <INPUT
                TYPE="text"
                SIZE="32"
                MAXLENGTH="32"
                NAME="attribute_value">
        </TD>
    </TR>
</TABLE>
<BR>
* Indicates a required field
<BR>
<BR>
<INPUT TYPE="submit" VALUE="Add Attribute Value">
</FORM>        

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
