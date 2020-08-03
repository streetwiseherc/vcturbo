<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   PRD_EDITBODY.ASP                                                       #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>

<%
Dim quantity
Dim monogram
Dim sku
Dim rsAttributeIds


quantity = Request("quantity")
if quantity = "" then    
   quantity = 1
end if

monogram = Request("monogram")

dim attrs(5)
sku = Request("sku")

attrs(0) = 1
attrs(1) = 1
attrs(2) = 1
attrs(3) = 1
attrs(4) = 1
if sku <> "" then
    Dim rsCurrentAttributes
    set rsCurrentAttributes = conn.Execute("select attribute0, attribute1, attribute2, attribute3, attribute4 from vcturbo_product_variant where sku = " & CStr(sku))
    if Not rsCurrentAttributes.EOF then
        attrs(0) = rsCurrentAttributes("attribute0").Value
        attrs(1) = rsCurrentAttributes("attribute1").Value
        attrs(2) = rsCurrentAttributes("attribute2").Value
        attrs(3) = rsCurrentAttributes("attribute3").Value
        attrs(4) = rsCurrentAttributes("attribute4").Value
    end if 
end if
%> 

<TABLE BORDER="1" CELLPADDING="2" CELLSPACING="3">
    <TR>
        <TD COLSPAN=2><IMG SRC="<% = pageURL("assets/images/options.gif") %>" ALIGN="center"></TD>
    </TR>
    <TR>
        <TD>Quantity:</TD>
        <TD><INPUT TYPE="TEXT" NAME="quantity" VALUE=<% = quantity %> SIZE="5" ALIGN="left"></TD>
    </TR>
    
    <% if Cint(rsProduct("monogramable").Value) = 1 then %>
        <TR>
            <TD>Monogram:</TD>
            <TD><INPUT TYPE="TEXT" NAME="monogram" VALUE="<% = monogram %>" MAXLENGTH="6" SIZE="7"></TD>
        </TR>
    <% end if %>    

    <% 
    REM Get all the attributes
    set rsAttributeIds = conn.Execute("select attribute_id, attribute_value from vcturbo_product_attribute  where pfid = '"& Replace(rsProduct("pfid").Value, "'", "''") &"' and attribute_index = 0 order by attribute_id ")

    While not rsAttributeIds.EOF %>
        <TR>
            <TD><% = rsAttributeIds("attribute_value").Value %>:</TD>
            <TD><% Dim test %>
                <% test = rsAttributeIds("attribute_id").Value %>
                <SELECT name=<% = "attr" & CStr(test) %>>
                    <% Dim rsAttributes
                    set rsAttributes = Conn.Execute("select pfid, attribute_id, attribute_index, attribute_value from vcturbo_product_attribute  where pfid = '"& Replace(rsProduct("pfid").Value, "'", "''") &"' and attribute_id = "& rsAttributeIds("attribute_id") &" and attribute_index > 0 order by attribute_index")

                    While not rsAttributes.EOF %>
                        <% = mscsPage.Option(rsAttributes("attribute_index").Value, attrs(test)) %> <% = rsAttributes("attribute_value").Value %>
                        <% rsAttributes.MoveNext
                    Wend %>
                </SELECT>
            </TD>
        </TR>
        <% rsAttributeIds.MoveNext
    Wend %>
</TABLE>
<P>
