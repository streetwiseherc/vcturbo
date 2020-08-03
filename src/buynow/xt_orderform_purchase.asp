<!--#INCLUDE FILE="include/shop.asp" -->
<!--#INCLUDE FILE="../include/util.asp" -->

<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   XT_ORDERFORM_PURCHASE.ASP                                               %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>
<%
function UtilGetBillingInformation(mscsOrderForm, errorList)
	REM - Get the new billing information
	mscsOrderForm.bill_to_name = mscsPage.RequestString("bill_to_name", null, 1, 50)
	if IsNull(mscsOrderForm.bill_to_name) then
	    errorList.Add("Bill to name must be between 1 and 50 characters")
	end if
	mscsOrderForm.bill_to_street  = mscsPage.RequestString("bill_to_street", null, 1, 50)
	if IsNull(mscsOrderForm.bill_to_street) then
	    errorList.Add("Bill to street must be between 1 and 50 characters")
	end if
	mscsOrderForm.bill_to_city = mscsPage.RequestString("bill_to_city", null, 1, 50)
	if IsNull(mscsOrderForm.bill_to_city) then
	    errorList.Add("Bill to city must be between 1 and 50 characters")
	end if
	mscsOrderForm.bill_to_country = mscsPage.RequestString("bill_to_country", null, 1, 50)
	if IsNull(mscsOrderForm.bill_to_country) then
	    errorList.Add("Bill to country must be between 1 and 50 characters")
	end if
	if mscsOrderForm.bill_to_country = "USA" then
	    mscsOrderForm.bill_to_state = mscsPage.RequestString("bill_to_state", null, 2, 2)
	    if IsNull(mscsOrderForm.bill_to_state) then
	        errorList.Add("Bill to state must be a two character code")
	    end if
	else
	    mscsOrderForm.bill_to_state = mscsPage.RequestString("bill_to_state", null, 1, 30)
	    if IsNull(mscsOrderForm.bill_to_state) then
	        errorList.Add("Bill to state must be between 1 and 30 characters")
	    end if
	end if
	mscsOrderForm.bill_to_zip = mscsPage.RequestString("bill_to_zip", null, 5, 15)
	if IsNull(mscsOrderForm.bill_to_zip) then
	    errorList.Add("Bill to zip must be between 5 and 15 characters")
	end if
	mscsOrderForm.bill_to_phone = mscsPage.RequestString("bill_to_phone", null, 8, 20)
	if IsNull(mscsOrderForm.bill_to_phone) then
	    errorList.Add("Bill to phone must be between 8 and 20 characters")
	end if
	mscsOrderForm.bill_to_email = mscsPage.RequestString("bill_to_email", null, 4, 50)
	if IsNull(mscsOrderForm.bill_to_email) then
	    errorList.Add("Bill to email must be between 4 and 20 characters")
	end if
	
	REM -- Payment method always credit
	mscsOrderForm.[_payment_method] = "credit"
	
	REM -- CC validation
	mscsOrderForm.cc_name = mscsPage.RequestString("cc_name", null, 1, 100)
	if IsNull(mscsOrderForm.cc_name) then
	    errorList.Add("Credit card name must be between 1 and 100 characters")
	end if
	mscsOrderForm.cc_type = mscsPage.RequestString("cc_type", null, 4, 20)
	if IsNull(mscsOrderForm.cc_type) or (mscsOrderForm.cc_type <> "Visa" and mscsOrderForm.cc_type <> "Mastercard") then
	    errorList.Add("Credit card type must be between 4 and 20 characters (one of 'Visa' or 'Mastercard')")
	end if
	mscsOrderForm.[_cc_number] = mscsPage.RequestString("_cc_number", null, 13, 19)
	if IsNull(mscsOrderForm.[_cc_number]) then
	    errorList.Add("Credit card number must be between 13 and 19 characters")
	end if
	mscsOrderForm.[_cc_expmonth] = mscsPage.RequestNumber("_cc_expmonth", null, 1, 12)
	if IsNull(mscsOrderForm.[_cc_expmonth]) then
	    errorList.Add("Expiration month must be a number between 1 and 12")
	end if
	mscsOrderForm.[_cc_expyear] = mscsPage.RequestNumber("_cc_expyear", null, 1996, 2003)
	if IsNull(mscsOrderForm.[_cc_expyear]) then
	    errorList.Add("Expiration year must be a number between 1996 and 2003")
	end if
end function

%>

<%
REM - Create a list to store errors
Dim errorList
set errorList = Server.CreateObject("Commerce.SimpleList")

REM First create the order form
Dim mscsOrderForm
set mscsOrderForm = Server.CreateObject("Commerce.OrderForm")

REM Set some order fields
mscsOrderForm.shopper_id = mscsShopperId

REM Get the item data into the order form
call UtilGetItemsCheck(mscsOrderForm, mscsShopperId)

REM Get the order level data
call UtilGetOrder(mscsOrderForm, mscsShopperId)

REM Get the defaults
call UtilGetDefaults(mscsOrderForm)

REM Fill in the billing information
call UtilGetBillingInformation(mscsOrderForm, errorList)

REM if their are any errors at this point, report them
if errorList.Count > 0 then
%>
    <!--#INCLUDE FILE="include/error.asp" -->
    <%Response.End
end if

REM - Process any verify with fields
call mscsPage.ProcessVerifyWith(mscsOrderForm)

REM Get the basic pipe context
Set mscsPipeContext = UtilGetPipeContext()
    
REM Create and run the pipe for computing the amounts/subtotal
Dim errorLevelBasket
errorLevelBasket = UtilRunPipe("basket.pcf", mscsOrderForm, mscsPipeContext)

REM Create and run the total pipe
Dim errorLevelTotal 
errorLevelTotal = UtilRunPipe("total.pcf", mscsOrderForm, mscsPipeContext)
Dim errorLevel
if errorLevelBasket > errorLevelTotal then
	errorLevel = errorLevelBasket
else
	errorLevel = errorLevelTotal
end if

REM -- If no errors, run the actual purchase
if errorList.Count = 0 and mscsOrderForm.[_Basket_Errors].Count = 0 and mscsOrderForm.[_Purchase_Errors].Count = 0 and errorLevel = 1 then
    Dim mscsPipeContext
	set mscsPipeContext = UtilGetPipeContext()
        
    REM Run the transacted pipe
    errorLevel = UtilRunTxPipe("purchase.pcf", mscsOrderForm, mscsPipeContext)
end if

REM Process errors from running the pipelines so they are displayed to the shopper
if mscsOrderForm.[_Purchase_Errors].Count > 0 or errorLevel > 1 then
	if mscsOrderForm.[_Basket_Errors].Count > 0 then        
		for each errorStr in mscsOrderForm.[_Basket_Errors]
			errorList.Add(errorStr)
		next
	end if

	if mscsOrderForm.[_Purchase_Errors].Count > 0 then
        for each errorStr in mscsOrderForm.[_Purchase_Errors]
            errorList.Add(errorStr)
        next
    end if

    REM Make sure if an error occured there is at least one error to display
    if errorList.Count = 0 then
    	errorList.Add("Unable to complete purchase at this time")
    end if
end if

if errorList.Count > 0 then
    Dim rsOrder
	set rsOrder = UtilGetOrderRowset(mscsShopperId)	
	
	rsOrder("bill_to_name").Value = mscsOrderForm.bill_to_name
	rsOrder("bill_to_street").Value = mscsOrderForm.bill_to_street
	rsOrder("bill_to_city").Value = mscsOrderForm.bill_to_city
	rsOrder("bill_to_state").Value = mscsOrderForm.bill_to_state
	rsOrder("bill_to_zip").Value = mscsOrderForm.bill_to_zip
	rsOrder("bill_to_country").Value = mscsOrderForm.bill_to_country
	rsOrder("bill_to_email").Value = mscsOrderForm.bill_to_email
	rsOrder("bill_to_phone").Value = mscsOrderForm.bill_to_phone
	rsOrder("cc_name").Value = mscsOrderForm.cc_name
	rsOrder.Update
%>

<%
else
    REM Purchase was successful....delete he items and the order from the database
    Dim sql
	sql = "delete from vcturbo_basket where shopper_id = '" + mscsShopperId + "'"
	conn.Execute(sql)

	sql = "delete from vcturbo_basket_item where shopper_id = '" + mscsShopperId + "'"
	conn.Execute(sql)
    
    Dim pageRedirect
    pageRedirect = "receipt.asp?" & mscsPage.URLShopperArgs("order_id", mscsOrderForm.order_id)
    call Response.Redirect(pageRedirect)
end if
%>


 


