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

<% pageTitle = "Cross Promotions" %>
<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<A HREF="promo-cross_new.asp"><H2> Add New Cross Promotion </H2></A>


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

<%
ListNoRowsText = "No Cross Promotions in table"
ListSQL = "select vcturbo_promo_cross.pfid, vcturbo_promo_cross.related_pfid, vcturbo_promo_cross.description, a.name name, b.name rname from vcturbo_promo_cross, vcturbo_product_family a, vcturbo_product_family b where vcturbo_promo_cross.pfid = a.pfid and vcturbo_promo_cross.related_pfid = b.pfid"
%>

<!--#INCLUDE FILE="include/list.asp" -->
<P>

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
