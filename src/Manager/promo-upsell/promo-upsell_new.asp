<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   PROMO-UPSELL_NEW.ASP                                                    %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "New Upsell" %>


<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<FORM METHOD="POST" ACTION="xt_data_add_update.asp">

<TABLE>
    <TR>
        <TD>Product:</TD>
        <TD>
            <SELECT NAME="pfid">
                <%
                cmdTemp.CommandText = "SELECT * FROM vcturbo_product_family order by pfid"
                Set rsProduct = Server.CreateObject("ADODB.Recordset")
                rsProduct.Open cmdTemp, , adOpenKeyset, adLockReadOnly
                If rsProduct.RecordCount > 0 Then 
                    While Not rsProduct.EOF %>
                        <% = mscsPage.Option(Cstr(rsProduct("pfid").Value), 1) %> <% = rsProduct("pfid").Value %>:<% = rsProduct("name").Value %>
                        <% rsProduct.MoveNext
                    Wend
                End If
                %>
            </SELECT>
        </TD>
    </TR>
    <TR>
        <TD>Related Product:</TD>
        <TD>
            <SELECT NAME="related_pfid">
                <%
                cmdTemp.CommandText = "SELECT * FROM vcturbo_product_family order by pfid"
                Set rsProduct = Server.CreateObject("ADODB.Recordset")
                rsProduct.Open cmdTemp, , adOpenKeyset, adLockReadOnly
                If rsProduct.RecordCount > 0 Then 
                    While Not rsProduct.EOF %>
                        <% = mscsPage.Option(Cstr(rsProduct("pfid").Value), 1) %> <% = rsProduct("pfid").Value %>:<% = rsProduct("name").Value %>
                        <% rsProduct.MoveNext
                    Wend
                End If
                %>
            </SELECT>
        </TD>
    </TR>
    <TR>
        <% REM label:  %>
        <TD VALIGN=TOP> 
            Slogan: 
        </TD>
        <% REM value:  %>
        <TD VALIGN=TOP> 
            <TEXTAREA COLS="40" 
                      ROWS="5" 
                      NAME="description" 
                      WRAP = "virtual">
            </TEXTAREA> 
        </TD>
    </TR>
</TABLE>
<BR>
<INPUT TYPE="submit" VALUE="Add Upsell">
</FORM>

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
