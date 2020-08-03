<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   PRODUCT_LIST.ASP                                                        %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "Products" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<A HREF="product_new.asp"> <H2>Add New Product</H2> </A>

<%
function ListRow() %>
        <TD VALIGN=TOP ALIGN=CENTER> <% = RowCount %>    </TD>
        <TD VALIGN=TOP ALIGN=LEFT  > <A HREF="<% = listElemTemplate & "?" & mscsPage.URLArgs("pfid", Cstr(rsList("pfid"))) %>"> <% = rsList("pfid") %> </A> </TD>
        <TD VALIGN=TOP ALIGN=LEFT  > <% = rsList("name") %>   </TD>
        <TD VALIGN=TOP ALIGN=RIGHT> <% = MSCSDataFunctions.Money(Cstr(rsList("list_price"))) %> </TD>

<% end function
function ListHeader() %>
        <TH ALIGN=LEFT VALIGN=BOTTOM> #          </TH>
        <TH ALIGN=LEFT VALIGN=BOTTOM> PFId       </TH>
        <TH ALIGN=LEFT VALIGN=BOTTOM> Name       </TH>
        <TH ALIGN=LEFT VALIGN=BOTTOM> Price      </TH>

<% end function
listElemTemplate = "product_edit.asp"
listElement = "products"
listNoRowsText = "No products in table"
ListSQL = "SELECT * FROM vcturbo_product_family ORDER BY pfid"
%>
<!--#INCLUDE FILE="include/list.asp" -->

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
