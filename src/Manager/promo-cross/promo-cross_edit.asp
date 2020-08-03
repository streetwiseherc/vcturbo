<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   PROMO-CROSS_EDIT.ASP                                                    %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "Edit Cross Promo '" &  Request("pfid") & "=>" & Request("rpfid") & "'" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<%
sql = "select vcturbo_promo_cross.pfid, vcturbo_promo_cross.related_pfid, vcturbo_promo_cross.description, a.name name, b.name rname from vcturbo_promo_cross, vcturbo_product_family a, vcturbo_product_family b where vcturbo_promo_cross.pfid = '" & Replace(Request("pfid"), "'", "''") & "' and related_pfid = '" & Replace(Request("rpfid"), "'", "''") & "' and vcturbo_promo_cross.pfid = a.pfid and vcturbo_promo_cross.related_pfid = b.pfid"
cmdTemp.CommandText = sql
Set rsCpromo = Server.CreateObject("ADODB.Recordset")
rsCpromo.Open cmdTemp, , adOpenStatic, adLockReadOnly
%>


<FORM METHOD="POST" ACTION="xt_data_add_update.asp">
<INPUT TYPE="HIDDEN" NAME="op"   VALUE="update">

<TABLE >
    <TR>
        <% REM label:  %>
        <TD VALIGN=TOP>
            PFId:Name
        </TD>

        <% REM value:  %>
        <TD VALIGN=TOP>
            <INPUT TYPE="HIDDEN" NAME="pfid" VALUE="<% = Request("pfid") %>">
            <STRONG><% = Request("pfid") %>:<% = rsCpromo("name").Value %></STRONG>
        </TD>
    </TR>
    <TR>
        <% REM label:  %>
        <TD VALIGN=TOP>
            Related PFId:Name
        </TD>

        <% REM value:  %>
        <TD VALIGN=TOP>
            <INPUT TYPE="HIDDEN" NAME="related_pfid" VALUE="<% = Request("rpfid") %>">
            <STRONG><% = Request("rpfid") %>:<% = rsCpromo("rname").Value %></STRONG>
        </TD>
    </TR>
    <TR>
        <% REM label:  %>
        <TD VALIGN=TOP>
            Slogan:
        </TD>

        <% REM value:  %>
        <TD VALIGN=TOP>
            <TEXTAREA
                COLS="40"
                ROWS="5"
                NAME="description"
                WRAP="virtual"><% = rsCpromo("description").Value %></TEXTAREA>
        </TD>
    </TR>
</TABLE>

<BR>
<TABLE>
<TR>
<TD>
    <INPUT TYPE="submit" VALUE="Update Cross Promo...">
    </FORM>
</TD>

<TD>

    <FORM METHOD="POST" ACTION="confirm.asp">
        <INPUT TYPE="HIDDEN" NAME="confirm_message" VALUE="Delete the cross promo '<% = Request("pfid") & "=>" & Request("rpfid") %>?">
        <INPUT TYPE="HIDDEN" NAME="confirm_action" VALUE="xt_data_delete.asp">
        <INPUT TYPE="HIDDEN" NAME="pfid"  VALUE="<% = Request("pfid") %>">
        <INPUT TYPE="HIDDEN" NAME="rpfid" VALUE="<% = Request("rpfid") %>">
        <INPUT TYPE="SUBMIT"              VALUE="Delete Cross Promo...">
    </FORM>
</TD>
</TR>
</TABLE>

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->

