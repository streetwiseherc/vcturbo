<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   PROMO-CROSS_LIST.ASP                                                    %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<!--#<% pageTitle = "Cross Promotions Target" %>
<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<A HREF="promo-cross_new.asp"><H2> Add New Cross Promotion </H2></A>

<%
ListNoRowsText = "No Cross Promotions in table"
ListSQL        = "select c.pfid, c.related_pfid, a.name name, b.name rname from vcturbo_promo_cross c, vcturbo_product_family b, vcturbo_product_family a "
ListSQL         = ListSQL + "where c.pfid = '" + Request("pfid") + "' and c.pfid = a.pfid and c.related_pfid = b.pfid"
%>

<% 
function ListHeader()
Response.Write("<TH ALIGN=LEFT VALIGN=BOTTOM> # </TH>")
Response.Write("<TH ALIGN=LEFT VALIGN=BOTTOM> From  </TH>")
Response.Write("<TH ALIGN=LEFT VALIGN=BOTTOM> To    </TH>")
end function

function ListRow()
%>
<TD VALIGN=TOP ALIGN=CENTER> 
<A HREF="<% = "promo-cross_edit.asp?" & mscsPage.URLArgs("rpfid", rsList("related_pfid").Value, "pfid", rsList("pfid").Value) %>"><% = RowCount %></A>
</TD>
<TD VALIGN=TOP>  <% = rsList("pfid").Value %>:<% = rsList("name").Value %> </A> </TD>
<TD VALIGN=TOP> <% = rsList("related_pfid").Value %>:<% = rsList("rname").Value %> </TD>
<%
end function
%>

<!--#INCLUDE FILE="include/list.asp" -->
<P>

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->

