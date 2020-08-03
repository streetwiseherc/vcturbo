<!--#INCLUDE FILE="include/shop.asp" -->
<!--#INCLUDE FILE="include/order.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   XT_ORDERFORM_EDITITEM.ASP                                              #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>

<% REM Create the error list %>
<OBJECT RUNAT=Server ID=errorList PROGID="Commerce.SimpleList"></OBJECT>

<%
REM Get the index of the item to delete
Dim item_id
item_id = mscsPage.RequestNumber("item_id")

Dim quantity
Dim pfid
Dim sql
Dim rsProduct
Dim monogram
Dim sku
Dim price
Dim name
Dim attr_text

if IsNull(item_id) then
	errorList.Add("Invalid line item")
else
    quantity = mscsPage.RequestNumber("quantity", null, 0, 999)

    if IsNull(quantity) then
        errorList.Add("Quantity must be a number between 1 and 999.")
    else
        if quantity = 0 then 
        	call OrderDeleteItem(mscsShopperId, item_id)
        else
            dim attrs(5)
            
            attrs(0) = mscsPage.RequestString("attr0", "0")
            attrs(1) = mscsPage.RequestString("attr1", "0")
            attrs(2) = mscsPage.RequestString("attr2", "0")
            attrs(3) = mscsPage.RequestString("attr3", "0")
            attrs(4) = mscsPage.RequestString("attr4", "0")
            pfid = Request("pfid")

            sql = OrderGetItemSQL(pfid, attrs)
                
            Set cmdTemp = Server.CreateObject("ADODB.Command")
            cmdTemp.CommandType = adCmdText
            Set cmdTemp.ActiveConnection = conn
            cmdTemp.CommandText = sql
            Set rsProduct = Server.CreateObject("ADODB.Recordset")
            rsProduct.Open cmdTemp, , adOpenKeyset, adLockReadOnly

            if rsProduct.recordcount = 0 then
                errorList.Add("This product is not available at this time.")
            else
                monogram = mscsPage.RequestString("monogram", "", 0, 3)
                if IsNull(monogram) then 
                    errorList.Add("Monogram must be at most 3 characters.")
                else
                   	sku = rsProduct("sku").Value
            	    price = rsProduct("list_price").Value
                    name = rsProduct("name").Value
					attr_text = OrderGetAttributeText(rsProduct, attrs)
                	call OrderUpdateItem(mscsShopperId, item_id, sku, pfid, quantity, price, attr_text, monogram, name)
                end if
            end if
        end if
    end if
end if

if errorList.Count > 0 then
%>
    <!--#INCLUDE FILE="include/error.asp" -->
<%
else    
    Dim pageRedirect
    pageRedirect = "basket.asp?" & mscsPage.URLShopperArgs()
    call Response.Redirect(pageRedirect)
end if
%>
