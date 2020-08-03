<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   ATTR_LIST.ASP                                                           %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "Product Attributes" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<%
cmdTemp.CommandText = "select count(*) c from vcturbo_product_attribute where pfid = '" & Replace(Request("pfid"), "'", "''") & "' and attribute_index = 0"
Set rsList = Server.CreateObject("ADODB.Recordset")
rsList.Open cmdTemp, , adOpenStatic, adLockReadOnly
%>
<A HREF="<% = "product_edit.asp?" & mscsPage.URLArgs("pfid", Cstr(Request("pfid"))) %>"><H2> PFId: <% = Request("pfid") %> </H2></A>
<%
if rsList("c").Value  < 5 then %>
    <A HREF="<% = "attr_new.asp?" & mscsPage.URLArgs("pfid", Cstr(Request("pfid")), "attribute_id", Cstr(rsList("c").Value)) %>"> <H2>Add New Attribute</H2> </A>
<% end if 
rsList.Close %>

<%
function ListRow() %>
        <TD VALIGN=TOP ALIGN=CENTER> <A HREF="<% = listElemTemplate & "?" & mscsPage.URLArgs("pfid", CStr(Request("pfid")), "attribute_id", CStr(rsList("attribute_id")), "attribute_name", CStr(rsList("attribute_value"))) %>"> <% = rsList("attribute_id") %> </A></TD>
        <FORM METHOD=POST ACTION="xt_data_add_update.asp">
            <INPUT TYPE="HIDDEN" NAME="goto"        VALUE="attr_list.asp">
            <INPUT TYPE="HIDDEN" NAME="op"          VALUE="update">
            <INPUT TYPE="HIDDEN" NAME="type"        VALUE="attribute">
            <INPUT TYPE=HIDDEN NAME="pfid"            VALUE="<% = Request("pfid") %>">
            <INPUT TYPE=HIDDEN NAME="attribute_id"    VALUE="<% = rsList("attribute_id") %>">
            <INPUT TYPE=HIDDEN NAME="attribute_index" VALUE="0">
            <TD VALIGN=TOP>
            <INPUT TYPE=TEXT NAME="attribute_value" MAXLENGTH="20" VALUE="<% = rsList("attribute_value") %>">
            </TD>
            <TD VALIGN=TOP><INPUT TYPE=SUBMIT VALUE=Update></TD>
        </FORM>
<% end function
function ListHeader() %>
        <TH ALIGN=LEFT VALIGN=BOTTOM> Id     </TH>
        <TH ALIGN=LEFT VALIGN=BOTTOM> Name   </TH>
        <TH ALIGN=LEFT VALIGN=BOTTOM> &nbsp; </TH>
        <TH ALIGN=LEFT VALIGN=BOTTOM> &nbsp; </TH>

<% end function
REM Prepare arguements for the list include
listElemTemplate = "attr_value_list.asp"
listElement = "attributes"
listNoRowsText = "No attributes in table"
ListSQL = "select attribute_id, attribute_value from vcturbo_product_attribute where pfid = '" & Replace(Request("pfid"), "'", "''") & "' and attribute_index = 0 order by attribute_id"
%>
<!--#INCLUDE FILE="include/list.asp" -->

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
