<!--#INCLUDE FILE="include/shop.asp" -->
<!--#INCLUDE FILE="../include/util.asp" -->
<!--#INCLUDE FILE="../include/order.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   XT_ORDERFORM_PREPARE.ASP                                                %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<%
function OrderFormPrepareArgs(byRef rsOrder, byRef errorList)
    OrderFormPrepareArgs = true

    REM -- Validate the ship to address
    rsOrder("ship_to_name").Value = mscsPage.RequestString("ship_to_name", null, 1, 50)
    if IsNull(rsOrder("ship_to_name").Value) then
        errorList.Add("Ship to name must be between 1 and 50 characters")
    end if
    rsOrder("ship_to_street").Value  = mscsPage.RequestString("ship_to_street", null, 1, 50)
    if IsNull(rsOrder("ship_to_street").Value) then
        errorList.Add("Ship to street must be between 1 and 50 characters")
    end if
    rsOrder("ship_to_city").Value = mscsPage.RequestString("ship_to_city", null, 1, 50)
    if IsNull(rsOrder("ship_to_city").Value) then
        errorList.Add("Ship to city must be between 1 and 50 characters")
    end if
    rsOrder("ship_to_country").Value = mscsPage.RequestString("ship_to_country", null, 1, 50)
    if IsNull(rsOrder("ship_to_country").Value) then
        errorList.Add("Ship to country must be between 1 and 50 characters")
    end if
    if rsOrder("ship_to_country").Value = "USA" then
        rsOrder("ship_to_state").Value = mscsPage.RequestString("ship_to_state", null, 2, 2)
        if IsNull(rsOrder("ship_to_state").Value) then
            errorList.Add("Ship to state must be a two character code")
        end if
    else
        rsOrder("ship_to_state").Value = mscsPage.RequestString("ship_to_state", null, 1, 30)
        if IsNull(rsOrder("ship_to_state").Value) then
            errorList.Add("Ship to state must be between 1 and 30 characters")
        end if
    end if
    rsOrder("ship_to_zip").Value = mscsPage.RequestString("ship_to_zip", null, 5, 15)
    if IsNull(rsOrder("ship_to_zip").Value) then
        errorList.Add("Ship to zip must be between 5 and 15 characters")
    end if
    rsOrder("ship_to_phone").Value = mscsPage.RequestString("ship_to_phone", null, 8, 20)
    if IsNull(rsOrder("ship_to_phone").Value) then
        errorList.Add("Ship to phone must be between 8 and 20 characters")
    end if
    rsOrder("ship_to_email").Value = mscsPage.RequestString("ship_to_email", null, 1, 50)
    if IsNull(rsOrder("ship_to_email").Value) then
        errorList.Add("Ship to email must be between 1 and 50 characters")
    end if

end function
%>

<%
REM Create a list to store errors
Dim errorList
set errorList = Server.CreateObject("Commerce.SimpleList")

REM Get the rowset representing the order
Dim rsOrder
Set rsOrder = UtilGetOrderRowset(mscsShopperId)

REM Retreive the args from the form into the rowset
call OrderFormPrepareArgs(rsOrder, errorList)

REM If errors, report them. Otherwise save to the database and redirect
if errorList.Count >0 then
%>
    <!--#INCLUDE FILE="include/error.asp" -->
<%

else
	rsOrder.Update
    Dim pageRedirect
    pageRedirect = "payment.asp?" & mscsPage.URLShopperArgs("use_form", mscsPage.RequestString("use_form", "0"))
    
    call Response.Redirect(pageRedirect)
end if

%>