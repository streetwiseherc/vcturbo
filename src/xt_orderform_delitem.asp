<!--#INCLUDE FILE="include/shop.asp" -->
<!--#INCLUDE FILE="include/order.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   XT_ORDERFORM_DELITEM.ASP                                               #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>

<% REM Create the error list %>
<OBJECT RUNAT=Server ID=errorList PROGID="Commerce.SimpleList"></OBJECT>

<%
Dim item_id
item_id = mscsPage.RequestNumber("item_id")
if IsNull(item_id) then
	errorList.Add("Invalid line item")
%>
	<!--#INCLUDE FILE="include/error.asp" -->
<%
else
	call OrderDeleteItem(mscsShopperId, item_id)
    Dim pageRedirect

	pageRedirect = "basket.asp?" & mscsPage.URLShopperArgs()

	call Response.Redirect(pageRedirect)
end if

%>
