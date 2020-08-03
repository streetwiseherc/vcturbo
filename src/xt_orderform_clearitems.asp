<!--#INCLUDE FILE="include/shop.asp" -->
<!--#INCLUDE FILE="include/order.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   XT_ORDERFORM_CLEARITEMS.ASP                                            #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>

<%
call OrderClearItems(mscsShopperId)
Dim pageRedirect 
pageRedirect = "basket.asp?" & mscsPage.URLShopperArgs()

call Response.Redirect(pageRedirect)
%>
 
