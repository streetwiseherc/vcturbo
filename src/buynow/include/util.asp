<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   UTIL.ASP                                                                %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<%

function UtilGetPipeContext()
    Dim pipeContext
    Set pipeContext = Server.CreateObject("Commerce.Dictionary")
    Set pipeContext("MessageManager")       = MSCSMessageManager
    Set pipeContext("DataFunctions")        = MSCSDataFunctions
    Set pipeContext("QueryMap")             = MSCSQueryMap
    Set pipeContext("ConnectionStringMap")  = MSCSSite.ConnectionStringMap
    pipeContext("DefaultConnectionString")  = MSCSSite.DefaultConnectionString
    pipeContext("Language")                 = "usa"

    Set UtilGetPipeContext = pipeContext
end function

function UtilRunPipe(file, orderForm, pipeContext)
    Dim pipeline
    Dim errorLevel

    set pipeline = Server.CreateObject("Commerce.MtsPipeline")
    
    call pipeline.LoadPipe(Request.ServerVariables("APPL_PHYSICAL_PATH") + "\config\" + file)

    REM Call pipeline.SetLogFile("c:\temp\pipeline.log")

    errorLevel = pipeline.Execute(1, orderForm, pipeContext, 0)

    UtilRunPipe = errorLevel
end function

function UtilRunTxPipe(file, orderForm, pipeContext)
    Dim Pipeline
    Dim errorLevel

    set pipeline = Server.CreateObject("Commerce.MtsTxPipeline")
    
    call pipeline.LoadPipe(Request.ServerVariables("APPL_PHYSICAL_PATH") + "\config\" + file)

    REM Call pipeline.SetLogFile("c:\temp\pipeline.log")

    errorLevel = pipeline.Execute(1, orderForm, pipeContext, 0)

    UtilRunTxPipe = errorLevel
end function

function UtilGetItemsCheck(mscsOrderForm, mscsShopperId)
    Dim rsOrderItem
    Dim delete
    Dim price
    Dim item
    Dim sql
    
    sql = "select vcturbo_product_variant.sku, vcturbo_product_family.pfid, quantity, monogram, item_id, vcturbo_basket_item.name, list_price, placed_price, attr_text "
    sql = sql + "from vcturbo_basket_item, vcturbo_product_family, vcturbo_product_variant where "
    sql = sql + "shopper_id = '" &  mscsShopperId & "' and vcturbo_basket_item.pfid *= vcturbo_product_family.pfid and vcturbo_product_variant.sku =* vcturbo_basket_item.sku "
    sql = sql + "order by item_id"
    set rsOrderItem = conn.Execute(sql)

  	delete = 0
    price = 0
    while Not rsOrderItem.EOF
		if IsNull(rsOrderItem("sku").Value) then
			delete = 0
		else
			if rsOrderItem("list_price").Value <> rsOrderItem("placed_price").Value then
				price = 1
			end if

            set item = mscsOrderForm.AddItem(rsOrderItem("sku").Value, rsOrderItem("quantity").Value, 0)
            
	    	item.monogram = rsOrderItem("monogram").Value
    		item.attr_text = rsOrderItem("attr_text").Value
    		item.item_id = rsOrderItem("item_id").Value
    		item.pfid = rsOrderItem("pfid").Value
    		item.[_product_list_price] = rsOrderItem("list_price").Value
    		item.name = rsOrderItem("name").Value
    	end if

		rsOrderItem.MoveNext
	wend

	if delete or price then
   		if delete then
   			sql = "delete from vcturbo_basket_item where shopper_id = '" + mscsShopperId + "' and item_id not in (select item_id from vcturbo_basket_item, vcturbo_product_family, vcturbo_product_variant where shopper_id = '" &  mscsShopperId & "' and vcturbo_basket_item.pfid = vcturbo_product_family.pfid and vcturbo_product_variant.sku = vcturbo_basket_item.sku)"
   			call conn.Execute(sql)
			mscsOrderForm.[_basket_errors].Add("Items in your basket had to be deleted because they no longer exist in this store")
		end if
		if price then
			sql = "update vcturbo_basket_item set placed_price = list_price from vcturbo_product_variant b, vcturbo_product_family c where b.pfid = c.pfid and b.sku = vcturbo_basket_item.sku and shopper_id = '" & mscsShopperId & "'"
			conn.Execute(sql)
			mscsOrderForm.[_basket_errors].Add("Prices of items in your basket have changed")
		end if
    end if


end function

function UtilGetItems(mscsOrderForm, mscsShopperId)
    Dim rsOrderItem
    Dim sql

	REM Get the item data
	sql = "select sku, item.pfid, item.quantity, item.monogram, item.item_id, item.name, item.placed_price, attr_text "
	sql = sql + "from vcturbo_basket_item item "
	sql = sql + "where shopper_id = '" & mscsShopperId & "' "
	sql = sql + "order by item_id"
    set rsOrderItem = conn.Execute(sql)

    while Not rsOrderItem.EOF
        Dim item
        set item = mscsOrderForm.AddItem(rsOrderItem("sku").Value, rsOrderItem("quantity").Value, 0)
        
		item.monogram = rsOrderItem("monogram").Value
		item.attr_text = rsOrderItem("attr_text").Value
		item.item_id = rsOrderItem("item_id").Value
		item.pfid = rsOrderItem("pfid").Value
		item.name = rsOrderItem("name").Value
		item.[_product_list_price] = rsOrderItem("placed_price").Value

		rsOrderItem.MoveNext
	wend
end function

function UtilGetDefaults(mscsOrderForm)
	if IsNull(mscsOrderForm.ship_to_name) or IsEmpty(mscsOrderForm.ship_to_name) or mscsOrderForm.ship_to_name = "" then
		mscsOrderForm.ship_to_name = mscsOrderForm.[_shopper_name]
		mscsOrderForm.ship_to_street = mscsOrderForm.[_shopper_street]
		mscsOrderForm.ship_to_city = mscsOrderForm.[_shopper_city]
		mscsOrderForm.ship_to_state = mscsOrderForm.[_shopper_state]
		mscsOrderForm.ship_to_zip = mscsOrderForm.[_shopper_zip]
		mscsOrderForm.ship_to_country = mscsOrderForm.[_shopper_country]
		mscsOrderForm.ship_to_email = mscsOrderForm.[_shopper_email]
		mscsOrderForm.ship_to_phone = mscsOrderForm.[_shopper_phone]
	end if

	if IsNull(mscsOrderForm.bill_to_name) or IsEmpty(mscsOrderForm.bill_to_name) or mscsOrderForm.bill_to_name = "" then
		mscsOrderForm.bill_to_name = mscsOrderForm.[_shopper_name]
		mscsOrderForm.bill_to_street = mscsOrderForm.[_shopper_street]
		mscsOrderForm.bill_to_city = mscsOrderForm.[_shopper_city]
		mscsOrderForm.bill_to_state = mscsOrderForm.[_shopper_state]
		mscsOrderForm.bill_to_zip = mscsOrderForm.[_shopper_zip]
		mscsOrderForm.bill_to_country = mscsOrderForm.[_shopper_country]
		mscsOrderForm.bill_to_email = mscsOrderForm.[_shopper_email]
		mscsOrderForm.bill_to_phone = mscsOrderForm.[_shopper_phone]
	end if

	if IsNull(mscsOrderForm.cc_name) or IsEmpty(mscsOrderForm.cc_name) or mscsOrderForm.cc_name = "" then
		mscsOrderForm.cc_name = mscsOrderForm.[_shopper_name]
	end if
end function

function UtilGetOrder(mscsOrderForm, mscsShopperId)
    Dim sql
    Dim rsOrder

	REM Select the current order level data
	sql = "select shopper_id, order_id, created, modified, ship_to_name, ship_to_street, ship_to_city, ship_to_state, ship_to_zip, ship_to_country, ship_to_phone, ship_to_email, bill_to_name, bill_to_street, bill_to_city, bill_to_state, bill_to_zip, bill_to_country, bill_to_phone, bill_to_email, cc_name from vcturbo_basket where shopper_id = '" & mscsShopperId & "'"

	set rsOrder = conn.Execute(sql)

	if Not(rsOrder.EOF) then
		mscsOrderForm.order_id = rsOrder("order_id").Value

		mscsOrderForm.ship_to_name = rsOrder("ship_to_name").Value
		mscsOrderForm.ship_to_street = rsOrder("ship_to_street").Value
		mscsOrderForm.ship_to_city = rsOrder("ship_to_city").Value
		mscsOrderForm.ship_to_state = rsOrder("ship_to_state").Value
		mscsOrderForm.ship_to_zip = rsOrder("ship_to_zip").Value
		mscsOrderForm.ship_to_country = rsOrder("ship_to_country").Value
		mscsOrderForm.ship_to_email = rsOrder("ship_to_email").Value
		mscsOrderForm.ship_to_phone = rsOrder("ship_to_phone").Value

		mscsOrderForm.bill_to_name = rsOrder("bill_to_name").Value
		mscsOrderForm.bill_to_street = rsOrder("bill_to_street").Value
		mscsOrderForm.bill_to_city = rsOrder("bill_to_city").Value
		mscsOrderForm.bill_to_state = rsOrder("bill_to_state").Value
		mscsOrderForm.bill_to_zip = rsOrder("bill_to_zip").Value
		mscsOrderForm.bill_to_country = rsOrder("bill_to_country").Value
		mscsOrderForm.bill_to_email = rsOrder("bill_to_email").Value
		mscsOrderForm.bill_to_phone = rsOrder("bill_to_phone").Value

		mscsOrderForm.cc_name = mscsOrderForm.[_shopper_name]		
	end if

end function

function UtilGetOrderRowset(mscsShopperId)
    Dim sql
	REM Select the current order level data
	sql = "select shopper_id, order_id, created, modified, ship_to_name, ship_to_street, ship_to_city, ship_to_state, ship_to_zip, ship_to_country, ship_to_phone, ship_to_email, bill_to_name, bill_to_street, bill_to_city, bill_to_state, bill_to_zip, bill_to_country, bill_to_phone, bill_to_email, cc_name from vcturbo_basket where shopper_id = '" & mscsShopperId & "'"

	REM Create a record set for the order level data
    Dim rsOrder
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	cmdTemp.CommandType = adCmdText
	cmdTemp.CommandText = sql
	Set cmdTemp.ActiveConnection = conn
	Set rsOrder = Server.CreateObject("ADODB.RecordSet")
	rsOrder.Open cmdTemp, , adOpenKeyset, adLockOptimistic

	REM If there is no record, then create one
	if rsOrder.EOF then
		rsOrder.AddNew
		rsOrder("shopper_id").Value = mscsShopperId
		rsOrder("created").Value = now
		rsOrder("modified").Value = now
	else
		rsOrder("modified").Value = now
	end if	

	Set UtilGetOrderRowset = rsOrder
end function

function UtilRunBasket(mscsShopperId)
    REM Create the order form
    Dim mscsOrderForm
    Dim mscsPipeContext
    Dim errorLevel

    set mscsOrderForm = Server.CreateObject("Commerce.OrderForm")

    REM Set some order fields
    mscsOrderForm.shopper_id = mscsShopperId

	REM Get the items
    call UtilGetItems(mscsOrderForm, mscsShopperId)

    REM Get the pipe context
    Set mscsPipeContext = UtilGetPipeContext()    

    REM Create and run the pipe
    errorLevel = UtilRunPipe("basket.pcf", mscsOrderForm, mscsPipeContext)
	
	REM Return the orderform
    Set UtilRunBasket = mscsOrderForm
    
end function

function UtilRunPlan(mscsShopperId)
    Dim mscsPipeContext
    Dim mscsOrderForm
    Dim errorLevel

    REM Create the order form
    set mscsOrderForm = Server.CreateObject("Commerce.OrderForm")

    REM Set some order fields
    mscsOrderForm.shopper_id = mscsShopperId

	REM Get the items
    call UtilGetItemsCheck(mscsOrderForm, mscsShopperId)

    REM Get the pipe context
    Set mscsPipeContext = UtilGetPipeContext()    

    REM Create and run the pipe
    errorLevel = UtilRunPipe("basket.pcf", mscsOrderForm, mscsPipeContext)

    REM Get the order information
    call UtilGetOrder(mscsOrderForm, mscsShopperId)

    REM Not get defaults for the bill/pay information
    call UtilGetDefaults(mscsOrderForm)
	
	REM - Process any verify with fields
	call mscsPage.ProcessVerifyWith(mscsOrderForm)

	REM Create and run the total pipe
	errorLevel = UtilRunPipe("total.pcf", mscsOrderForm, mscsPipeContext)
	
	REM Return the order form
	set UtilRunPlan = mscsOrderForm
end function
%>

    

    

