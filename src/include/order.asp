<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   ORDER.ASP  				                                            #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>
<%

function OrderGetReceipt(shopper_id, order_id)
    Dim sql
    Dim rsOrder	
    Dim rsOrderItem
    Dim mscsOrderForm

	sql = "select order_id, shopper_id, created, modified, ship_to_name, ship_to_street, ship_to_city, ship_to_state, ship_to_zip, ship_to_country, ship_to_phone, ship_to_email, bill_to_name, bill_to_street, bill_to_city, bill_to_state, bill_to_zip, bill_to_country, bill_to_phone, bill_to_email, oadjust_subtotal, shipping_total, handling_total, tax_total, tax_included, cc_name, cc_type, total_total from vcturbo_receipt "
	sql = sql + " where shopper_id = '" + shopper_id + "' and order_id = " + CStr(order_id)
	set rsOrder = conn.Execute(sql)

	sql = "select shopper_id, created, modified, item_id, sku, pfid, name, quantity, product_list_price, iadjust_regularprice, iadjust_currentprice, oadjust_adjustedprice, monogram, attr_text from vcturbo_receipt_item "
	sql = sql + " where shopper_id = '" + shopper_id + "' and order_id = " + CStr(order_id)
	set rsOrderItem = conn.Execute(sql)

	set mscsOrderForm = Server.CreateObject("Commerce.OrderForm")
	
	if Not(rsOrder.EOF) then
		mscsOrderForm.shopper_id = shopper_id
		mscsOrderForm.order_id   = order_id

		mscsOrderForm.created  = rsOrder("created").Value
		mscsOrderForm.modified = rsOrder("modified").Value
		
		mscsOrderForm.ship_to_name    = rsOrder("ship_to_name").Value
		mscsOrderForm.ship_to_street  = rsOrder("ship_to_street").Value
		mscsOrderForm.ship_to_city    = rsOrder("ship_to_city").Value
		mscsOrderForm.ship_to_state   = rsOrder("ship_to_state").Value
		mscsOrderForm.ship_to_zip     = rsOrder("ship_to_zip").Value
		mscsOrderForm.ship_to_country = rsOrder("ship_to_country").Value
		mscsOrderForm.ship_to_email   = rsOrder("ship_to_email").Value
		mscsOrderForm.ship_to_phone   = rsOrder("ship_to_phone").Value

		mscsOrderForm.bill_to_name    = rsOrder("bill_to_name").Value
		mscsOrderForm.bill_to_street  = rsOrder("bill_to_street").Value
		mscsOrderForm.bill_to_city    = rsOrder("bill_to_city").Value
		mscsOrderForm.bill_to_state   = rsOrder("bill_to_state").Value
		mscsOrderForm.bill_to_zip     = rsOrder("bill_to_zip").Value
		mscsOrderForm.bill_to_country = rsOrder("bill_to_country").Value
		mscsOrderForm.bill_to_email   = rsOrder("bill_to_email").Value
		mscsOrderForm.bill_to_phone   = rsOrder("bill_to_phone").Value

		mscsOrderForm.[_oadjust_subtotal] = rsOrder("oadjust_subtotal").Value
		mscsOrderForm.[_shipping_total]   = rsOrder("shipping_total").Value
        mscsOrderForm.[_handling_total]   = rsOrder("handling_total").Value
        mscsOrderForm.[_tax_total]        = rsOrder("tax_total").Value
        mscsOrderForm.[_tax_included]     = rsOrder("tax_included").Value
        mscsOrderForm.[_cc_name]          = rsOrder("cc_name").Value
        mscsOrderForm.[_cc_type]          = rsOrder("cc_type").Value
        
        mscsOrderForm.[_total_total] = rsOrder("total_total").Value

        Dim fldSKU                  
        Dim fldQuantity             
        Dim fldPFID                 
        Dim fldMonogram             
        Dim fldAttr_Text            
        Dim fldName                 
        Dim fldProduct_List_Price  
        Dim fldIAdjust_RegularPrice 
        Dim fldIAdjust_CurrentPrice 
        Dim fldOAdjust_AdjustedPrice 
         
        Set fldSKU                  = rsOrderItem("SKU")
        Set fldQuantity             = rsOrderItem("Quantity")
        Set fldPFID                 = rsOrderItem("PFID")
        Set fldMonogram             = rsOrderItem("Monogram")
        Set fldAttr_Text            = rsOrderItem("Attr_Text")
        Set fldName                 = rsOrderItem("Name")
        Set fldProduct_List_Price   = rsOrderItem("Product_List_Price")
        Set fldIAdjust_RegularPrice = rsOrderItem("IAdjust_RegularPrice")
        Set fldIAdjust_CurrentPrice = rsOrderItem("IAdjust_CurrentPrice")
        Set fldOAdjust_AdjustedPrice = rsOrderItem("OAdjust_AdjustedPrice")

        Do Until rsOrderItem.EOF
            Dim Item
		    Set Item = mscsOrderForm.AddItem(fldSKU, fldQuantity, 0)

	        Item.pfid = fldPFID
	        Item.monogram = fldMonogram
	        Item.attr_text = fldAttr_Text
	        Item.name = fldName
	        Item.[_product_list_price] = fldProduct_List_Price
	        Item.[_iadjust_regularprice] = fldIAdjust_RegularPrice
	        Item.[_iadjust_currentprice] = fldIAdjust_CurrentPrice
	        Item.[_oadjust_adjustedprice] = fldOAdjust_AdjustedPrice

		    rsOrderItem.MoveNext

        Loop

	end if
	
	set OrderGetReceipt = mscsOrderForm
end function

function OrderClearItems(shopper_id)
    Dim sql	
	sql = "delete from vcturbo_basket_item where shopper_id = '" + shopper_id + "'"

	conn.Execute(sql)
end function

function OrderDeleteItem(shopper_id, item_id)
    Dim sql	
	sql = "delete from vcturbo_basket_item where shopper_id = '" + shopper_id + "' and item_id = " + CStr(item_id)

	conn.Execute(sql)
end function

function OrderDelete(shopper_id)
	
	sql = "delete from vcturbo_basket where shopper_id = '" + mscsShopperId + "'"
	conn.Execute(sql)

	sql = "delete from vcturbo_basket_item where shopper_id = '" + mscsShopperId + "'"
	conn.Execute(sql)
end function


function OrderAddItem(shopper_id, sku, pfid, quantity, placed_price, attr_text, monogram, name)

	REM sql = "select max(item_id) value from vcturbo_basket_item where shopper_id = '" + shopper_id +"'"
	REM set rsData = conn.Execute(sql, nRows, adCmdText)

	REM cmdTemp.CommandText = sql
	
	REM value = rsData("value").Value
	REM if (IsNull(value)) then
	REM 	item_id = 0
	REM else
	REM 	item_id = value + 1
	REM end if
	
	sql = "insert into vcturbo_basket_item "
	sql = sql + "(shopper_id, created, modified, sku, pfid, quantity, placed_price, attr_text, monogram, name)"
	sql = sql + " values ( '" + shopper_id + "', GetDate(), GetDate(), " + CStr(sku) + ", '" + pfid + "', " + CStr(quantity) + ", " + CStr(placed_price) + ", '" + Replace(attr_text, "'", "''") + "', '" + Replace(monogram, "'", "''") + "', '" + Replace(name, "'", "''") + "')"

	conn.Execute(sql)	
end function

function OrderUpdateItem(shopper_id, item_id, sku, pfid, quantity, placed_price, attr_text, monogram, name)
	
	sql = "update vcturbo_basket_item set "
	sql = sql + " sku = " + CStr(sku)
	sql = sql + ", modified = GetDate()"
	sql = sql + ", pfid = '" + pfid + "'"
	sql = sql + ", quantity = " + CStr(quantity)
	sql = sql + ", placed_price = " + CStr(placed_price)
	sql = sql + ", attr_text = '" + Replace(attr_text, "'", "''") + "'"
	sql = sql + ", monogram = '" + Replace(monogram, "'", "''") + "'"
	sql = sql + ", name = '" + Replace(name, "'", "''") + "'"
	sql = sql + " where shopper_id = '" + shopper_id + "' and item_id = " + CStr(item_id)
	
	conn.Execute(sql)
end function

function OrderDeleteOrder(shopper_id)
	
	sql = "delete from vcturbo_basket where shopper_id = '" + shopper_id + "'"
	conn.Execute(sql)

	sql = "delete from vcturbo_basket_item where shopper_id = '" + shopper_id + "'"
	
	conn.Execute(sql)
end function

function OrderGetItemSQL(pfid, attrs)
    Dim qpfid
    Dim i
    Dim s
    Dim f
    Dim w
    Dim t

    qpfid = "'" + Replace(pfid, "'", "''") + "'"
    
    s = "b.name name, a.sku sku, b.list_price list_price"
    f = "vcturbo_product_variant a, vcturbo_product_family b"
    w = "a.pfid = b.pfid and a.pfid = " & qpfid

    i = 0
    while i < 5 
        if attrs(i) <> 0 then
            t = "t" + CStr(i)
            s = s + ", " + t + ".attribute_value attr" + CStr(i)
            f = f + ", vcturbo_product_attribute " + t
            w = w + " and a.attribute" + CStr(i) + " = " + CStr(attrs(i)) + " and " + t + ".attribute_index = " + CStr(attrs(i)) + " and " + t + ".attribute_id = " + CStr(i) + " and " + t + ".pfid = " + qpfid
        else
            w = w + " and a.attribute" + CStr(i) + " = 0"
        end if

        i = i + 1
    wend

    OrderGetItemSQL = "select " + s + " from " + f + " where " + w
end function

function OrderGetAttributeText(rsProduct, attrs)
    Dim attr_text 
    Dim i
    i = 0
    attr_text = ""
    while i < 5 
        if attrs(i) <> 0 then 
            if attr_text <> "" then
                attr_text = attr_text + ", "
            end if
            attr_text = attr_text + CStr(rsProduct("attr" + CStr(i)).Value)
        end if
        i = i + 1
    wend 

    OrderGetAttributeText = attr_text
end function



function OrderBasketAttrText(row_orderitem)
    Dim result
    Dim attr_text
    Dim monogram

    result = ""

    if IsNull(row_orderitem.monogram) or row_orderitem.monogram = " " then
        monogram=""
    else
        monogram=Cstr(row_orderitem.monogram)
    end if
    
    if IsNull(row_orderitem.attr_text) or row_orderitem.attr_text = " " then
        attr_text = ""
    else
        attr_text = row_orderitem.attr_text
    end if

   if monogram <> "" and attr_text <> "" then
        result = mscsPage.HTMLEncode(row_orderitem.attr_text) + ", Monogram: " + monogram 
   elseif monogram <> "" then
        result = "Monogram: " + monogram
   elseif attr_text <> "" then 
        result = attr_text
   end if

   OrderBasketAttrText = result
end function %>


