<!--#INCLUDE FILE = "include/shop.asp" -->
<!--#INCLUDE FILE = "include/util.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   XT_SHOPPER_NEW.ASP                                                     #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>

<%
function ShopperNew(byRef errorList)
    ShopperNew = NULL

    REM -- Get and validate name, street, city, state, zip, country, phone, email:
    Dim Shopper_name
    shopper_name = mscsPage.RequestString("name", null,1,50)
    if isnull(shopper_name) then
        errorList.Add("Name must be a string between 1 and 50 characters.")
    end if
    Dim shopper_street
    shopper_street = mscsPage.RequestString("street", null,1,50)
    if isnull(shopper_street) then
        errorList.Add("Street must be a string between 1 and 50 characters.")
    end if
    Dim shopper_city
    shopper_city = mscsPage.RequestString("city", null,1,50)
    if isnull(shopper_city) then
        errorList.Add("City must be a string between 1 and 50 characters.")
    end if
    Dim shopper_country
    shopper_country = mscsPage.RequestString("country", null,1,50)
    if isnull(shopper_country) then
        errorList.Add("Country must be a string between 1 and 50 characters.")
    end if
    if shopper_country = "USA" then
        Dim shopper_state
        shopper_state = mscsPage.RequestString("state", null,2,2)
        if isnull(shopper_state) then
            errorList.Add("State must be 2 characters.")
        end if
    else
        shopper_state = mscsPage.RequestString("state", null, 1, 30)
        if isnull(shopper_state) then 
            errorList.Add("State must be between 1 and 30 characters.")
        end if
    end if
    Dim shopper_zip
    shopper_zip = mscsPage.RequestString("zip", null,5,15)
    if isnull(shopper_zip) then
        errorList.Add("Zip must be a string between 5 and 15 characters.")
    end if
    Dim shopper_phone
    shopper_phone = mscsPage.RequestString("phone", null,8,20)
    if isnull(shopper_phone) then
        errorList.Add("Phone must be a string between 8 and 16 characters.")
    end if
    Dim shopper_email
    shopper_email = mscsPage.RequestString("email", null,1,50)
    if isnull(shopper_email) then
        errorList.Add("E-Mail must be a string between 1 and 50 characters.")
    end if

    REM -- validate password:
    Dim shopper_pwd1
    Dim shopper_pwd2
    shopper_pwd1 = mscsPage.RequestString("pwd1", null,4,20)
    shopper_pwd2 = mscsPage.RequestString("pwd2", null,4,20)
    if isnull(shopper_pwd1) or isnull(shopper_pwd2) or shopper_pwd1 <> shopper_pwd2 then
        errorList.Add("Passwords must match and be between 4 and 20 characters.")
    end if

    if errorList.Count > 0 then
        exit function
    end if
    Dim shopper_id
    shopper_id = MSCSShopperManager.CreateShopperId()
    Dim sql
    sql = "insert into vcturbo_shopper (shopper_id, created, name, password, street, city, state, zip, country, phone, email) values ("
	sql = sql + "'" + shopper_id + "'"
	sql = sql + ",GetDate()"
	sql = sql + ",'" + Replace(shopper_name, "'", "''") + "'"
	sql = sql + ",'" + Replace(shopper_pwd1, "'", "''") + "'"
	sql = sql + ",'" + Replace(shopper_street, "'", "''") + "'"
	sql = sql + ",'" + Replace(shopper_city, "'", "''") + "'"
	sql = sql + ",'" + Replace(shopper_state, "'", "''") + "'"
	sql = sql + ",'" + Replace(shopper_zip, "'", "''") + "'"
	sql = sql + ",'" + Replace(shopper_country, "'", "''") + "'"
	sql = sql + ",'" + Replace(shopper_phone, "'", "''") + "'"
	sql = sql + ",'" + Replace(shopper_email, "'", "''") + "'"
	sql = sql + ")"

	on error resume next
	conn.Execute(sql)
	if Err <> 0 then
		if Err = -2147217873 then
            errorList.Add("Shopper could not be registered. A shopper with this e-mail is already registered.")
        else
            errorList.Add("Shopper could not be registered at this time.")
        end if 
    end if
    on error goto 0
     
    ShopperNew = shopper_id
end function
%>

<OBJECT RUNAT=Server ID=errorList PROGID="Commerce.SimpleList"></OBJECT>

<%
Dim shopper_id
shopper_id =  ShopperNew(errorList)
if errorList.Count > 0 then
%>
    <!--#INCLUDE FILE="include/error.asp" -->
<%
else
    call mscsPage.PutShopperId(shopper_id)
    Dim pageRedirect
    pageRedirect = "home.asp?" & mscsPage.URLShopperArgs()
    call Response.Redirect(pageRedirect)
end if 
%>
