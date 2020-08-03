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

<% pageTitle = "Product Attribute Values" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->


<%
cmdTemp.CommandText = "select attribute_id, attribute_value from vcturbo_product_attribute where pfid = '" & Replace(Request("pfid"),"'", "''") & "' and attribute_id = " & Request("attribute_id") & " and attribute_index = 0 order by attribute_id"
Set rsAname = Server.CreateObject("ADODB.Recordset")
rsAname.Open cmdTemp, , adOpenStatic, adLockReadOnly
%>
<%
cmdTemp.CommandText = "select * from vcturbo_product_attribute where pfid = '" & Replace(Request("pfid"),"'", "''") & "' and attribute_id = " & Request("attribute_id") & " and attribute_index > 0 order by attribute_index"
Set rsList = Server.CreateObject("ADODB.Recordset")
rsList.Open cmdTemp, , adOpenStatic, adLockReadOnly
%>
<A HREF="<% = "product_edit.asp?" & mscsPage.URLArgs("pfid", Cstr(Request("pfid"))) %>"><H2> PFId: <% = Request("pfid") %> </H2></A>
<A HREF="<% = "attr_list.asp?" & mscsPage.URLArgs("pfid", Cstr(Request("pfid"))) %>"><H2>Attribute List</H2></A>
<H1> Attribute Name: <% = rsAname("attribute_value") %> </H1>
<A HREF="<% = "attr_value_new.asp?" & mscsPage.URLArgs("pfid", Cstr(Request("pfid")), "attribute_id", Cstr(Request("attribute_id")), "attr_name", Cstr(rsAname("attribute_value").Value), "attribute_index", Clng(rsList.RecordCount + 1)) %>"> <H2>Add New Attribute Value</H2> </A>
<% 
rsAname.Close
rsList.Close
%>

<%
function ListRow() %>
        <TD VALIGN=TOP> <% = rsList("attribute_index") %> </TD>
        <FORM METHOD=POST ACTION="xt_data_add_update.asp">
            <INPUT TYPE="HIDDEN" NAME="goto"            VALUE="attr_value_list.asp">
            <INPUT TYPE="HIDDEN" NAME="op"              VALUE="update">
            <INPUT TYPE="HIDDEN" NAME="type"            VALUE="attribute value">
            <INPUT TYPE=HIDDEN   NAME=pfid              VALUE=<% = Request("pfid") %>>
            <INPUT TYPE=HIDDEN   NAME="attribute_id"    VALUE=<% = Request("attribute_id") %>>
            <INPUT TYPE=HIDDEN   NAME="attribute_index" VALUE=<% = rsList("attribute_index") %>>
            <TD VALIGN=TOP>
                <INPUT TYPE=TEXT NAME="attribute_value" MAXLENGTH="20" VALUE=<% = rsList("attribute_value") %>>
            </TD>
            <TD VALIGN=TOP><INPUT TYPE=SUBMIT VALUE=Update></TD>
        </FORM>
<% end function
function ListHeader() %>
        <TH ALIGN=LEFT VALIGN=BOTTOM> Id     </TH>
        <TH ALIGN=LEFT VALIGN=BOTTOM> Value  </TH>
        <TH ALIGN=LEFT VALIGN=BOTTOM> &nbsp; </TH>
        <TH ALIGN=LEFT VALIGN=BOTTOM> &nbsp; </TH>
<% end function
REM Prepare arguements for the list include
listElemTemplate = "variant_edit.asp"
listElement = "Attribute values"
listNoRowsText = "No attributes values in table"
ListSQL = "select * from vcturbo_product_attribute where pfid = '" & Replace(Request("pfid"),"'", "''") & "' and attribute_id = " & Request("attribute_id") & " and attribute_index > 0 order by attribute_index"
%>
<!--#INCLUDE FILE="include/list.asp" -->

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->

