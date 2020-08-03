<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   XT_DATA_ADD_UPDATE.ASP                                                  %>
<% REM   VC TURBO STORE                                                          %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<%
Set errorList = Server.CreateObject("Commerce.SimpleList")
promo_name = mscsPage.RequestString("promo_name", null, 1, 255)
if IsNull(promo_name) then
   errorList.Add("Promo name must be 1 and 255 characters")
end if
promo_type = mscsPage.RequestNumber("promo_type", null, 1, 100)
if IsNull(promo_type) then
   errorList.Add("Promo type must be a positive integer between 1 and 100")
end if
promo_description = mscsPage.RequestString("promo_description")
if Not(IsNull(promo_description)) and len(promo_description) > 2000 then
   errorList.Add("Description must be of less than 2000 characters")
end if
promo_rank = mscsPage.RequestNumber("promo_rank", null, 1, 100)
if IsNull(promo_rank) then
   errorList.Add("Rank must be a number between 1 and 100")
end if
active = mscsPage.RequestNumber("active", null, 0, 1)
if IsNull(active) then
   errorList.Add("Status must be 0 (off) or 1 (on)")
end if
date_start = mscsPage.RequestDate("date_start", "0")
if IsNull(date_start) then
   errorList.Add("Start date must be a valid date, (mm/dd/yy)")
end if
date_end = mscsPage.RequestDate("date_end", "0")
if IsNull(date_end) then
   errorList.Add("End date must be a valid date, (mm/dd/yy)")
end if

REM Shopper criteria
shopper_all = mscsPage.RequestNumber("shopper_all", null, 0, 1)
if IsNull(shopper_all) then
   errorList.Add("Eligible shoppers must be all, or specific shoppers")
end if
if shopper_all = 0 then
    shopper_column = mscsPage.RequestString("shopper_column", null, 1, 64)
    if IsNull(shopper_column) then
       errorList.Add("Eligible shoppers criteria must be between 1 and 64 characters")
    end if
    shopper_op = mscsPage.RequestString("shopper_op", null, 1, 2)
    if IsNull(shopper_op) or _
             (shopper_op <> "<" and shopper_op <> "<=" and _
               shopper_op <> "=" and shopper_op <> ">=" and _
               shopper_op <> ">" and shopper_op <> "<>" and _
                shopper_op <> "!=" and shopper_op <> "==") then
        errorList.Add("Eligible shoppers comparison condition invalid")
    end if
    shopper_value = mscsPage.RequestString("shopper_value", null, 1, 64)
    if IsNull(shopper_value) then
        errorList.Add("Eligible shoppers condition must be between 1 and 64 characters")
    end if
end if

REM Buy criteria
cond_all = mscsPage.RequestNumber("cond_all", null, 0, 1)
if IsNull(cond_all) then
    errorList.Add("Buy must be all, or specific products")
end if
if cond_all = 0 then
    cond_column = mscsPage.RequestString("cond_column", null, 1, 64)
    if IsNull(cond_column) then
        errorList.Add("Buy criteria must be between 1 and 64 characters")
    end if
    cond_op = mscsPage.RequestString("cond_op", null, 1, 2)
    if IsNull(cond_op) or _
        (cond_op <> "<" and cond_op <> "<=" and _
         cond_op <> "=" and cond_op <> ">=" and _
         cond_op <> ">" and cond_op <> "<>" and _
         cond_op <> "!=" and cond_op <> "==") then
        errorList.Add("Buy comparison condition invalid")
    end if
    cond_value = mscsPage.RequestString("cond_value", null, 1, 64)
    if IsNull(cond_value) then
        errorList.Add("Buy condition must be between 1 and 64 characters")
    end if
end if

cond_basis = mscsPage.RequestString("cond_basis", null, 1, 1)
if IsNull(cond_basis) or (cond_basis <> "Q" and cond_basis <> "P") then
    errorList.Add("Buy basis must be 'Q' (unit(s)) or 'P' ($ worth)")
end if
if cond_basis = "P" then
    cond_min = mscsPage.RequestMoneyAsNumber("cond_min", null, 0, 2147483647)
    if IsNull(cond_min) then
        errorList.Add("Buy amount must be a positive currency amount")
    end if
else
    cond_min = mscsPage.RequestNumber("cond_min", null, 1,  999)
    if IsNull(cond_min) then
        errorList.Add("Buy amount must be between 1 and 999")
    end if
end if

REM Award
award_all = mscsPage.RequestNumber("award_all", null, 0, 1)
if IsNull(award_all) then
    errorList.Add("Get must be all, or specific products")
end if
if award_all = 0 then
    award_column = mscsPage.RequestString("award_column", null, 1, 64)
    if IsNull(award_column) then
        errorList.Add("Get criteria must be between 1 and 64 characters")
    end if
    award_op = mscsPage.RequestString("award_op", null, 0, 2)
    if IsNull(award_op) or _
        (award_op <> "<" and award_op <> "<=" and _
         award_op <> "=" and award_op <> ">=" and _
         award_op <> ">" and award_op <> "<>" and _
         award_op <> "!=" and award_op <> "==") then
        errorList.Add("Get comaprison condition invalid")
    end if
    award_value = mscsPage.RequestString("award_value", null, 1, 64)
    if IsNull(award_value) then
        errorList.Add("Get condition must be between 1 and 64 characters")
    end if
end if

award_max = mscsPage.RequestNumber("award_max", null, 1, 999)
if IsNull(award_max) then
    errorList.Add("Get amount must be between 1 and 999")
end if

disjoint_cond_award = mscsPage.RequestNumber("disjoint_cond_award", 1, 0, 1)
if IsNull(disjoint_cond_award) then
    errorList.Add("Condition/Award relation invalid")
end if

disc_type = mscsPage.RequestString("disc_type", null, 1, 1)
if IsNull(disc_type) or (disc_type <> "$" and disc_type <> "%") then
    errorList.Add("At discount type must be '$' ($ off) or '%' (% off)")
else
    if disc_type = "$" then
        disc_value = mscsPage.RequestMoneyAsNumber("disc_value", null, 0, 2147483647)
        if IsNull(disc_value) then
            errorList.Add("At amount must be a positive dollar amount")
        end if
    else
        disc_value = mscsPage.RequestNumber("disc_value", null, 1, 100)
        if IsNull(disc_value) then
            errorList.Add("At percentage off must be between 1 and 100")
        end if
    end if
end if


if errorList.count = 0 then
    cmdTemp.CommandText = "SELECT * FROM vcturbo_promo_price WHERE promo_name =  '" & Replace(promo_name, "'", "''") & "'"

    Set recordSet = Server.CreateObject("ADODB.Recordset")
    recordSet.Open cmdTemp, , adOpenKeyset, adLockOptimistic

    op = mscsPage.RequestString("op", "add")

    if (op = "add" and recordSet.RecordCount = 0) or (op = "update" and recordSet.RecordCount = 1) then

        if op = "add" then
            recordSet.AddNew
        end if

        if op = "add" then
           recordSet("promo_name").Value = promo_name
        end if
        recordSet("promo_type").Value        = promo_type
        recordSet("promo_description").Value = promo_description
        recordSet("promo_rank").Value        = promo_rank
        recordSet("active").Value            = active
        recordSet("date_start").Value        = date_start
        recordSet("date_end").Value          = date_end
        recordSet("shopper_all").Value       = shopper_all
        recordSet("shopper_column").Value    = shopper_column
        recordSet("shopper_op").Value        = shopper_op
        recordSet("shopper_value").Value     = shopper_value
        recordSet("cond_all").Value          = cond_all
        recordSet("cond_column").Value       = cond_column
        recordSet("cond_op").Value           = cond_op
        recordSet("cond_value").Value        = cond_value
        recordSet("cond_basis").Value        = cond_basis
        recordSet("cond_min").Value          = cond_min
        recordSet("award_all").Value         = award_all
        recordSet("award_column").Value      = award_column
        recordSet("award_op").Value          = award_op
        recordSet("award_value").Value       = award_value
        recordSet("award_max").Value         = award_max
        recordSet("disjoint_cond_award")     = disjoint_cond_award
        recordSet("disc_value").Value        = disc_value
        recordSet("disc_type").Value         = disc_type

        recordSet.Update

        Response.Redirect "promo-price_list.asp"
    else
        if op = "add" then
            errorList.Add("This price promo already exists in the price promo table.")
        else
            errorList.Add("This price promo does not exist in the price promo table.")
        end if

    end if
end if
if errorList.Count <> 0 then
%>
    <!--#INCLUDE FILE="include/error.asp" -->
<%
end if
%>
 
