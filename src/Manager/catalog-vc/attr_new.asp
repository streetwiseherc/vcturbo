<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   ATTR_NEW.ASP                                                            %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "New Attribute" %>
<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<TABLE>
    <FORM METHOD="POST" ACTION="xt_data_add_update.asp">
        <INPUT TYPE="HIDDEN" NAME="type"         VALUE="attribute">
        <INPUT TYPE="HIDDEN" NAME="goto"         VALUE="attr_list.asp">
        <INPUT TYPE="HIDDEN" NAME="attribute_id" VALUE="<% = Request("attribute_id") %>">
        <INPUT TYPE="HIDDEN" NAME="pfid"         VALUE="<% = Request("pfid") %>">
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
                NAME="attribute_value"
                MAXLENGTH="20"
                VALUE= "">
        </TD>
    </TR>
</TABLE>
<BR>
* Indicates a required field
<BR>
<BR>
<INPUT TYPE="submit" VALUE="Add Attribute">
</FORM>        

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
