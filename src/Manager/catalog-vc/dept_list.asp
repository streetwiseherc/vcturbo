<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   DEPT_LIST.ASP                                                           %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "Departments" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<A HREF="dept_new.asp"> <H2>Add New Department</H2> </A>

<%
function ListRow() %>
        <TD VALIGN=TOP ALIGN=CENTER> <% = rsList("dept_id") %> </TD>
        <TD VALIGN=TOP ALIGN=CENTER> <% = rsList("parent_id") %> </TD>
        <TD VALIGN=TOP ALIGN=LEFT  > <A HREF="<% = listElemTemplate & "?" & mscsPage.URLArgs("dept_id", Cstr(rsList("dept_id")), "name", Cstr(rsList("name"))) %>"> <% = rsList("name") %></A> </TD>

<% end function
function ListHeader() %>
        <TH ALIGN=LEFT VALIGN=BOTTOM> Id             </TH>
        <TH ALIGN=LEFT VALIGN=BOTTOM> Parent Dept Id </TH>
        <TH ALIGN=LEFT VALIGN=BOTTOM> Name           </TH>

<% end function
REM Prepare arguements for the list include
listElemTemplate = "dept_edit.asp"
listElement = "departments"
listNoRowsText = "No departments in table"
ListSQL = "SELECT * FROM vcturbo_dept ORDER BY dept_id"
%>
<!--#INCLUDE FILE="include/list.asp" -->

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->


