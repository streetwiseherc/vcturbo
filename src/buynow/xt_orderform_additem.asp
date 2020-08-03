<!--#INCLUDE FILE="include/shop.asp" -->
<!--#INCLUDE FILE="../include/util.asp" -->
<!--#INCLUDE FILE="../include/order.asp" -->

<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   XT_ORDERFORM_ADDITEM.ASP                                               #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>

<%
Dim errorList
set errorList = Server.CreateObject("Commerce.SimpleList")

REM -- retrieve quantity:
Dim quantity
quantity = mscsPage.RequestNumber("quantity" , null , 1 , 999)
if isnull(quantity) then 
    errorList.Add("Quantity must be a number between 1 and 999.")
end if

REM -- add monogram to product in orderform
Dim monogram
monogram = mscsPage.RequestString("monogram", "", 0, 3)
if IsNull(monogram) then 
    errorList.Add("Monogram must be at most 3 characters.")
end if

dim attrs(5)

REM -- retreive product sku
attrs(0) = mscsPage.RequestString("attr0", "0")
attrs(1) = mscsPage.RequestString("attr1", "0")
attrs(2) = mscsPage.RequestString("attr2", "0")
attrs(3) = mscsPage.RequestString("attr3", "0")
attrs(4) = mscsPage.RequestString("attr4", "0")

Dim pfid
pfid = Request("pfid")

Dim sql
sql = OrderGetItemSQL(pfid, attrs)

    
Set cmdTemp = Server.CreateObject("ADODB.Command")
cmdTemp.CommandType = adCmdText

Set cmdTemp.ActiveConnection = conn
cmdTemp.CommandText = sql

Dim rsProduct
Set rsProduct = Server.CreateObject("ADODB.Recordset")
rsProduct.Open cmdTemp, , adOpenKeyset, adLockReadOnly 
   
if rsProduct.EOF then
    errorList.Add("This product is not available at this time.")
end if

if errorList.Count > 0 then
%>
    <!--#INCLUDE FILE="include/error.asp" -->
<%
else
    Dim attr_text
    Dim pfReplace
    Dim Sku
    Dim Price
    Dim name
    
    attr_text = OrderGetAttributeText(rsProduct, attrs)
    pfReplace = Replace(pfid, "'", "''")
    sku       = rsProduct("sku").Value
   	price     = rsProduct("list_price").Value
   	name      = rsProduct("name").Value
   	call OrderAddItem(mscsShopperID, sku, pfid, quantity, price, attr_text, monogram, name)

    Dim rsCount
   	set rsCount = conn.Execute("select count(*) c from vcturbo_promo_upsell a, vcturbo_promo_cross b where a.pfid = '" & pfReplace & "' or b.pfid = '" & pfReplace & "'")

   	if rsCount("c").Value > 0 then
       	Response.Redirect "added.asp?" & mscsPage.URLShopperArgs()
    else
        Response.Redirect "shipping.asp?" & mscsPage.URLShopperArgs()
    end if
end if %>
