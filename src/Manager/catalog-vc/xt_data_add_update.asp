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
Select Case Request.Form("type") 
    Case "product" 
        pfid = mscsPage.RequestString("pfid", null, 1, 30)
        if IsNull(pfid) then
            errorList.Add("PFid must be between 1 and 30 characters.")
        end if
        dept_id = mscsPage.RequestNumber("dept_id")
        if IsNull(dept_id) then
            errorList.Add("Dept must be a number.")
        end if
        monogramable = mscsPage.RequestNumber("monogramable", 0, 0, 1)
        if IsNull(monogramable) then
            errorList.Add("Monogramable must be a number.")
        end if
        name = mscsPage.RequestString("name",null,1,255)
        if IsNull(name) then
            errorList.Add("Name must be 1 and 255 characters.")
        end if
        short_description = mscsPage.RequestString("short_description")
        if len(short_description) > 255 then
            errorList.Add("Short description must be less than 255 characters.")
        end if
        long_description = mscsPage.RequestString("long_description")
        image_filename = mscsPage.RequestString("image_filename")
        if len(image_filename) > 255 then
            errorList.Add("Image file must be less than 255 characters.")
        end if
        list_price = mscsPage.RequestMoneyAsNumber("list_price", null, 0, 2147483647)
        if IsNull(list_price) then
            errorList.Add("Price must be a number greater than or equal to 0.")
        end if
    Case "variant"
        sku = mscsPage.RequestNumber("sku")
        if IsNull(sku) then
            errorList.Add("Sku must be a number.")
        end if
        pfid = mscsPage.RequestString("pfid", null, 1, 30)
        if IsNull(pfid) then
            errorList.Add("PFid must be between 1 and 30 characters.")
        end if
        attribute0 = mscsPage.RequestNumber("attribute0",null)
        if IsNull(attribute0) then
            errorList.Add("Attribute 0 must be a number.")
        end if
        attribute1 = mscsPage.RequestNumber("attribute1",null)
        if IsNull(attribute1) then
            errorList.Add("Attribute 1 must be a number.")
        end if
        attribute2 = mscsPage.RequestNumber("attribute2",null)
        if IsNull(attribute2)then
            errorList.Add("Attribute 2 must be a number.")
        end if
        attribute3 = mscsPage.RequestNumber("attribute3",null)
        if IsNull(attribute3) then
            errorList.Add("Attribute 3 must be a number.")
        end if
        attribute4 = mscsPage.RequestNumber("attribute4",null)
        if IsNull(attribute4) then
            errorList.Add("Attribute 4 must be a number.")
        end if
    Case "attribute"
        pfid = mscsPage.RequestString("pfid", null, 1, 30)
        if IsNull(pfid) then
            errorList.Add("PFid must be between 1 and 30 characters.")
        end if
        attribute_id = mscsPage.RequestNumber("attribute_id",null,0,4)
        if IsNull(attribute_id) then
            errorList.Add("Id must be a number between 0 and 4.")
        end if
        attribute_value = mscsPage.RequestString("attribute_value",null,1,20)
        if IsNull(attribute_value) then
            errorList.Add("Name must be 1 and 20 characters.")
        end if
    Case "attribute value"
        pfid = mscsPage.RequestString("pfid", null, 1, 30)
        if IsNull(pfid) then
            errorList.Add("PFid must be between 1 and 30 characters.")
        end if
        attribute_id = mscsPage.RequestNumber("attribute_id",null,0,4)
        if IsNull(attribute_id) then
            errorList.Add("Attribute id must be a number between 0 and 4.")
        end if
        attribute_index = mscsPage.RequestNumber("attribute_index")
        if IsNull(attribute_index) then
            errorList.Add("Attribute index must be a number.")
        end if
        attribute_value = mscsPage.RequestString("attribute_value",null,1,20)
        if IsNull(attribute_value) then
            errorList.Add("Attribute value must be 1 and 20 characters.")
        end if
    Case "department"
        dept_id = mscsPage.RequestNumber("dept_id")
        if IsNull(dept_id) or (dept_id <= 0) then
            errorList.Add("ID must be a number and must be greater than 0")
        end if
        parent_id = mscsPage.RequestNumber("parent_id")
        if IsNull(parent_id) then
            errorList.Add("Parent ID must be a number")
        end if
        name = mscsPage.RequestString("name", null, 1, 255)
        if IsNull(name) then
            errorList.Add("Name must be between 1 and 255 characters.")
        end if
        description = mscsPage.RequestString("description", null)
        if Not(IsNull(description)) and len(description) > 2000 then
            errorList.Add("Description must be less than 2000 characters.")
        end if
    
End Select

if errorList.count = 0 then
    Select Case Request.Form("type") 
        Case "product" 
            cmdTemp.CommandText = "SELECT * FROM vcturbo_product_family WHERE pfid = '" & Replace(pfid, "'", "''") & "'"
        Case "variant"
            cmdTemp.CommandText = "SELECT * FROM vcturbo_product_variant WHERE (sku = " & sku & ") or ((pfid = '" & Replace(pfid, "'", "''") & "' and attribute0 = " & attribute0 & " and attribute1 = " & attribute1 & " and attribute2 = " & attribute2 & " and attribute3 = " & attribute3 & " and attribute4 = " & attribute4 & "))"
       Case "attribute"
           cmdTemp.CommandText = "SELECT * FROM vcturbo_product_attribute WHERE attribute_index = 0 and (attribute_id = " & attribute_id & " or attribute_value = '" & Replace(attribute_value, "'", "''") &"') and pfid = '" & Replace(pfid, "'", "''") & "'"
       Case "attribute value"
           cmdTemp.CommandText = "SELECT * FROM vcturbo_product_attribute WHERE attribute_id = " & attribute_id & " and (attribute_index = " & attribute_index & " or attribute_value = '" & Replace(attribute_value, "'", "''") &"') and pfid = '" & Replace(pfid, "'", "''" & "'") & "'"
        Case "department"
            cmdTemp.CommandText = "SELECT * FROM vcturbo_dept WHERE dept_id = " & dept_id 
    End Select

    Set recordSet = Server.CreateObject("ADODB.Recordset")
    recordSet.Open cmdTemp, , adOpenKeyset, adLockOptimistic

    op = mscsPage.RequestString("op", "add")

    if (op = "add" and recordSet.RecordCount = 0) or (op = "update" and recordSet.RecordCount = 1) then

        if op = "add" then
            recordSet.AddNew
        end if

        Select Case Request.Form("type") 
            Case "product" 
                if op = "add" then
                    recordSet("pfid").Value = pfid
                end if
                recordSet("list_price").Value        = list_price
                recordSet("dept_id").Value           = dept_id
                recordSet("name").Value              = name
                recordSet("short_description").Value = short_description
                recordSet("long_description").Value  = long_description
                recordSet("image_filename").Value    = image_filename
                recordSet("manufacturer_id").Value   = 1
                recordSet("intro_date").Value        = Date
                recordSet("date_changed").Value      = Now
                recordSet("monogramable").Value      = monogramable
            Case "variant"
                if op = "add" then
                    recordSet("sku").Value  = sku
                end if
                recordSet("pfid").Value  = pfid
                recordSet("attribute0").Value  = attribute0
                recordSet("attribute1").Value  = attribute1
                recordSet("attribute2").Value  = attribute2
                recordSet("attribute3").Value  = attribute3
                recordSet("attribute4").Value  = attribute4
            Case "attribute"
                if op = "add" then
                    recordSet("pfid").Value  =  pfid
                    recordSet("attribute_id").Value = attribute_id
                    recordSet("attribute_index").Value  = 0
                end if
                recordSet("attribute_value").Value  = attribute_value

            Case "attribute value"
                if op = "add" then
                    recordSet("pfid").Value  =  pfid
                    recordSet("attribute_id").Value = attribute_id
                    recordSet("attribute_index").Value  = attribute_index
                end if
                recordSet("attribute_value").Value  = attribute_value
            Case "department"
                if op = "add" then
                    recordSet("dept_id").Value = dept_id
                end if
                recordSet("parent_id").Value  = parent_id
                recordSet("name").Value =  name
                recordSet("date_changed").Value = Now
                recordSet("description").Value  = description
        End Select 

        recordSet.Update

        If Request.Form("pfid").Count = 0 Then
            Response.redirect Cstr(Request.Form("goto"))
        else
            If Request.Form("attribute_id").Count = 0 Then
                Response.redirect Cstr(Request.Form("goto")) & "?" & mscsPage.URLArgs( "pfid", CStr(Request.Form("pfid")))
            else
                Response.redirect Cstr(Request.Form("goto")) & "?" & mscsPage.URLArgs("pfid", CStr(Request.Form("pfid")), "attribute_id", CStr(Request.Form("attribute_id")))
            end if
        End If
    else
        if op = "add" then
            Select Case Request.Form("type") 
                Case "product" 
                    errorList.Add("This pfid already exists in the product family table.")
                Case "variant"
                    errorList.Add("This sku/attribute value set already exists in the product variant table.")
                Case "attribute"
                    errorList.Add("This attribute already exists in the attribute table.")
                Case "attribute value"
                    errorList.Add("This attribute value already exists in the attribute table.")
                Case "department"
                    errorList.Add("This department already exists in the department table.")
            End Select
        else
            Select Case Request.Form("type") 
                Case "product" 
                    errorList.Add("This product already exists, or this is an invalid pfid")
                Case "variant"
                    errorList.Add("This attribute combination already exists, or this is an invalid sku.")
                Case "attribute"
                    errorList.Add("This attribute already exists, or this is an invalid attribute.")
                Case "attribute value"
                    errorList.Add("This attribute value already exists, or this is an invalid attribute id.")
                Case "department"
                    errorList.Add("This department does not exist in the department table.")
            End Select
        end if

    end if
end if
if errorList.Count <> 0 then
%>
    <!--#INCLUDE FILE="include/error.asp" -->
<%
end if
%>
 
