<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   PROMO-CROSS_NEW.ASP                                                     %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "New Cross Promo" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<%
pfid = mscsPage.RequestString("pfid", "")
related_pfid = mscsPage.RequestString("related_pfid", "")
name = mscsPage.RequestString("name", "")
rname = mscsPage.RequestString("rname", "")
if name = "" and rname = "" then
    description = ""
else
    description = name + " -> " + rname
end if
%>


<FORM METHOD="POST"
    ACTION="xt_data_add_update.asp">
    <INPUT TYPE="HIDDEN" NAME="type" VALUE="promo-cross">

<TABLE >
    <TR>
        <TD> Product: </TD>
        <TD>
            <SELECT NAME="pfid">
                <%
                set rsProduct = conn.Execute("SELECT * FROM vcturbo_product_family order by pfid")
                While Not rsProduct.EOF %>
                    <% = mscsPage.Option(Cstr(rsProduct("pfid").Value), pfid) %> <% = rsProduct("pfid").Value %>:<% = rsProduct("name").Value %>
                    <% rsProduct.MoveNext
                Wend
                %>
            </SELECT>
        </TD>
    </TR>

    <TR>
        <TD> Related Product: </TD>
        <TD>
            <SELECT NAME="related_pfid">
                <%
                set rsProduct = conn.Execute("SELECT * FROM vcturbo_product_family order by pfid")
                While Not rsProduct.EOF %>
                    <% = mscsPage.Option(Cstr(rsProduct("pfid").Value), related_pfid) %> <% = rsProduct("pfid").Value %>:<% = rsProduct("name").Value %>
                    <% rsProduct.MoveNext
                Wend
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
                      WRAP="virtual"><% = description%></TEXTAREA> 
        </TD>
    </TR>
</TABLE>

<BR>
<INPUT TYPE="submit" VALUE="Add Cross Promo">

</FORM>

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
