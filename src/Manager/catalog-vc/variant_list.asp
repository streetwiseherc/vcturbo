<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   VARIANT_LIST.ASP                                                        %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "Product Variants" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<A HREF="<% = "product_edit.asp?" & mscsPage.URLArgs("pfid", Cstr(Request("pfid"))) %>"><H2>PFId: <% = mscsPage.HTMLEncode(Request("pfid")) %> </H2></A>
<A HREF="<% = "variant_new.asp?" & mscsPage.URLArgs("pfid", Cstr(Request("pfid"))) %>"><H2>Add New Product Variant</H2> </A>

<%
function ListRow() %>
        <%
        attr0 = rsList("attribute0")
        attr1 = rsList("attribute1")
        attr2 = rsList("attribute2")
        attr3 = rsList("attribute3")
        attr4 = rsList("attribute4")
        if CBool(attr0 <> "0") then
            cmdTemp.CommandText = "select attribute_value from vcturbo_product_attribute where pfid = '" & Replace(rsList("pfid").Value, "'", "''") & "' and attribute_id = 0 and attribute_index = " & attr0
            Set rsAttr = Server.CreateObject("ADODB.Recordset")
            rsAttr.Open cmdTemp, , adOpenKeyset, adLockReadOnly
            attr0 = rsAttr("attribute_value")
            rsAttr.Close
        end if
        if CBool(attr1 <> "0") then
            cmdTemp.CommandText = "select attribute_value from vcturbo_product_attribute where pfid = '" & Replace(rsList("pfid").Value, "'", "''") & "' and attribute_id = 1 and attribute_index = " & attr1
            Set rsAttr = Server.CreateObject("ADODB.Recordset")
            rsAttr.Open cmdTemp, , adOpenKeyset, adLockReadOnly
            attr1 = rsAttr("attribute_value")
            rsAttr.Close
        end if
        if CBool(attr2 <> "0") then
            cmdTemp.CommandText = "select attribute_value from vcturbo_product_attribute where pfid = '" & Replace(rsList("pfid").Value, "'", "''") & "' and attribute_id = 2 and attribute_index = " & attr2
            Set rsAttr = Server.CreateObject("ADODB.Recordset")
            rsAttr.Open cmdTemp, , adOpenKeyset, adLockReadOnly
            attr2 = rsAttr("attribute_value")
            rsAttr.Close
        end if
        if CBool(attr3 <> "0") then
            cmdTemp.CommandText = "select attribute_value from vcturbo_product_attribute where pfid = '" & Replace(rsList("pfid").Value, "'", "''") & "' and attribute_id = 3 and attribute_index = " & attr3
            Set rsAttr = Server.CreateObject("ADODB.Recordset")
            rsAttr.Open cmdTemp, , adOpenKeyset, adLockReadOnly
            attr3 = rsAttr("attribute_value")
            rsAttr.Close
        end if
        if CBool(attr4 <> "0") then
            cmdTemp.CommandText = "select attribute_value from vcturbo_product_attribute where pfid = '" & Replace(rsList("pfid").Value, "'", "''") & "' and attribute_id = 4 and attribute_index = " & attr4
            Set rsAttr = Server.CreateObject("ADODB.Recordset")
            rsAttr.Open cmdTemp, , adOpenKeyset, adLockReadOnly
            attr4 = rsAttr("attribute_value")
            rsAttr.Close
        end if
        %>
        <TD VALIGN=TOP> <% = RowCount %> </TD>
        <TD VALIGN=TOP ALIGN=CENTER>
            <A HREF="<% = listElemTemplate & "?" & mscsPage.URLArgs("sku", Cstr(rsList("sku"))) %>"> <% = rsList("sku") %> </A></TD>
        <TD VALIGN=TOP> <% if attr0 <> "0" then %> <% = attr0 %> <% end if %> </TD>
        <TD VALIGN=TOP> <% if attr1 <> "0" then %> <% = attr1 %> <% end if %> </TD>
        <TD VALIGN=TOP> <% if attr2 <> "0" then %> <% = attr2 %> <% end if %> </TD>
        <TD VALIGN=TOP> <% if attr3 <> "0" then %> <% = attr3 %> <% end if %> </TD>
        <TD VALIGN=TOP> <% if attr4 <> "0" then %> <% = attr4 %> <% end if %> </TD>
<% end function
function ListHeader() %>
        <TH ALIGN=CENTER VALIGN=BOTTOM> #   </TH>
        <TH ALIGN=LEFT   VALIGN=BOTTOM> Sku </TH>
        <%
        cmdTemp.CommandText = "select attribute_id, attribute_value from vcturbo_product_attribute where pfid = '" & Replace(Request("pfid"),"'", "''") & "' and attribute_index = 0 order by attribute_id"
        Set rsAttr = Server.CreateObject("ADODB.Recordset")
        rsAttr.Open cmdTemp, , adOpenKeyset, adLockReadOnly
        If rsAttr.RecordCount > 0 Then 
            While Not rsAttr.EOF %>
                <TH ALIGN=LEFT VALIGN=BOTTOM> <% = rsAttr("attribute_value").Value %> </TH>
                <% rsAttr.MoveNext
            Wend
        End If
        rsAttr.Close
        %>
<% end function
listElemTemplate = "variant_edit.asp"
listElement = "product variants"
listNoRowsText = "No product variants in table"
ListSQL = "select * from vcturbo_product_variant where pfid = '" & Replace(Request("pfid"),"'", "''") & "'"
%>
<!--#INCLUDE FILE="include/list.asp" -->

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
