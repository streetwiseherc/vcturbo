<!-- #INCLUDE FILE = "include/shop.asp" -->
<!-- #INCLUDE FILE = "include/util.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   XT_SHOPPER_LOOKUP.ASP                                                  #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>

<%
function ShopperLookup(byRef errorList)
    ShopperLookup = ""

    REM -- Validate the fields for the lookup
    Dim shopper_email
    shopper_email = mscsPage.RequestString("shopper_email", null,1,50)
    if isnull(shopper_email) then
        errorList.Add("E-Mail must be a string between 1 and 50")
    end if
    Dim shopper_password
    shopper_password = mscsPage.RequestString("shopper_password", null,4,20)
    if isnull(shopper_password) then
        errorList.Add("Password must be a string between 4 and 20")
    end if

    if errorList.Count > 0 then
        exit function
    end if

	REM Go against the database to find the shopper	
    Dim sql
    Dim rsShopper
    sql = "select shopper_id from vcturbo_shopper where email = '" & Replace(shopper_email, "'", "''") & "' and password = '" & Replace(shopper_password, "'", "''" & "'") & "'"
    set rsShopper = conn.Execute(sql)

	REM Return the error, or shopper id
    if rsShopper.EOF then
    	errorList.Add("Unable to find shopper e-mail/password.")
    	ShopperLookup = NULL
    else
    	ShopperLookup = rsShopper("shopper_id").Value
    end if
end function
%>

<%
REM -- set up vars:
Dim errorList
set errorList = Server.CreateObject("Commerce.SimpleList")

mscsShopperId = ShopperLookup(errorList)

if errorList.Count > 0 then
%>
    <!--#INCLUDE FILE="include/error.asp" -->
<%
else
    call mscsPage.PutShopperId(mscsShopperId)
    Dim pageRedirect
    pageRedirect = "home.asp?" & mscsPage.URLShopperArgs()
    
    call Response.Redirect(pageRedirect)
end if 
%>

