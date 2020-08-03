<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   BASKET_VIEW.ASP                                                         %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<%
shopper_id = Request("shopper_id")
%>

<% pageTitle = "Shopper '" &  mscsPage.HTMLEncode(shopper_id) & "'" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<br>
<A HREF="<% = Replace(mscsShopperPath, ":1", Request("shopper_id")) %>">View Shopper</A><br>
<br>

<%
    sql = "select sku, pfid, name, quantity, monogram, item_id, name, placed_price, attr_text "
    sql = sql + "from vcturbo_basket_item where "
    sql = sql + "shopper_id = '" &  shopper_id & "'"
    sql = sql + "order by item_id"
    set rsBasketItems = conn.Execute(sql)
    
    if Not(rsBasketItems.EOF) then %>
    <P>
    <TABLE>
        <% REM column headers:  %>
        <TR>
            <TH VALIGN=BOTTOM ALIGN=LEFT WIDTH=50>Sku</TH>
            <TH VALIGN=BOTTOM ALIGN=LEFT WIDTH=50>PFID</TH>
            <TH VALIGN=BOTTOM ALIGN=LEFT WIDTH=180>Item</TH>
            <TH VALIGN=BOTTOM ALIGN=CENTER></TH>
            <TH VALIGN=BOTTOM ALIGN=CENTER WIDTH=30>Qty</TH>
            <TH VALIGN=BOTTOM ALIGN=CENTER WIDTH=60>Price</TH>
        <TR>
        <TR> <TD COLSPAN=9 HEIGHT=1 BGCOLOR="black"></TD></TR>
        
        <% REM line items:  %>
        <% while not rsBasketItems.EOF %>
            <TR>
                <TD VALIGN=TOP ALIGN=LEFT>   <% = mscsPage.HTMLEncode(rsBasketItems("sku").Value) %></TD>
                <TD VALIGN=TOP ALIGN=LEFT>   <% = mscsPage.HTMLEncode(rsBasketItems("pfid").Value) %></TD>
                <TD VALIGN=TOP ALIGN=LEFT>   <% = mscsPage.HTMLEncode(rsBasketItems("name").Value) %></TD> 
                <TD VALIGN=TOP ALIGN=LEFT>   <% = mscsPage.HTMLEncode(rsBasketItems("monogram").Value) %> <% = mscsPage.Encode(rsBasketItems("attr_text").Value) %> </TD>
                <TD VALIGN=TOP ALIGN=CENTER> <% = mscsPage.HTMLEncode(rsBasketItems("quantity").Value) %></TD>
                <TD VALIGN=TOP ALIGN=RIGHT > <% = MSCSDataFunctions.Money(rsBasketItems("placed_price").Value) %></TD>
            </TR>
        <%  rsBasketItems.MoveNext
        wend %>
        <% REM divider:  %>
        <TR> <TD COLSPAN=10 HEIGHT=1 BGCOLOR="black"></TD></TR>
    
    
    </TABLE>
    <% else %>
    	<B>Basket is empty</B>
    <% end if %>

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
