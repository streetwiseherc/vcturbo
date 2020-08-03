<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   PROMO-UPSELL_LIST.ASP                                                   %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "Upsells" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<% REM   list vars: %>

<A HREF="promo-upsell_new.asp"><H2> Add New Upsell </H2></A>

<%
function ListRow() %>
    <TD VALIGN=TOP ALIGN=CENTER> 
    <A HREF="<% = "promo-upsell_edit.asp?" & mscsPage.URLArgs("rpfid", rsList("related_pfid").Value, "pfid", rsList("pfid").Value) %>"><% = RowCount %></A>
    </TD>
    <TD VALIGN=TOP><% = rsList("pfid") %>:<% = rsList("name").Value %> </A> </TD>
    <TD VALIGN=TOP><% = rsList("related_pfid") %>:<% = rsList("rname").Value %> </TD>
<% end function 
function ListHeader() %>
    <TH ALIGN=LEFT VALIGN=BOTTOM> #                 </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> PFId:Name         </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Related PFId:Name </TH>
<% end function 
ListNoRows = "No upsells found"
ListSQL = "select vcturbo_promo_upsell.pfid, vcturbo_promo_upsell.related_pfid, vcturbo_promo_upsell.description, a.name name, b.name rname from vcturbo_promo_upsell, vcturbo_product_family a, vcturbo_product_family b where vcturbo_promo_upsell.pfid = a.pfid and vcturbo_promo_upsell.related_pfid = b.pfid"
%>
<!--#INCLUDE FILE="include/list.asp" -->

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
