<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   DEFAULT.ASP                                                             %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<!--#Include File="include/SiteUtil.Asp"-->

<%
GetStatus Status, RevStatus
PCFFiles = GetPCFFiles
%>

<HTML>
<BODY BGCOLOR="<% = bgcolor %>" TEXT="<% = text %>" LINK="<% = color_link %>" ALINK="<% = color_alink %>" VLINK="<% = color_vlink %>">
<% pageTitle = Server.HTMLEncode(Application("MSCSSite").DisplayName) + " Site Manager" %>
<!--INCLUDE FILE="include/mgmt_header.asp" -->

<% REM   body: %>
<TABLE BORDER=1 CELLPADDING=5 CELLSPACING=3 BORDERCOLOR=black>
    <TR>
        <TD ALIGN=LEFT>
            <FONT FACE="Arial, sans-serif" SIZE="+2">M</FONT><FONT FACE="Arial, sans-serif" SIZE="+1">ERCHANDISING</FONT>
        </TD>
        <TD ALIGN=CENTER COLSPAN=3>
        <TABLE>
        <TR>
            <TD ALIGN=CENTER WIDTH=100>
                <A HREF="catalog-vc/dept_list.asp"><IMG SRC="MSCS_Images/admin/depts.gif" WIDTH=51 HEIGHT=41 BORDER=0 ALT="Departments"></A>
                <BR>
                <A HREF="catalog-vc/dept_list.asp"><FONT FACE="Arial, sans-serif">Departments</FONT></A>
            </TD>
            <TD ALIGN=CENTER WIDTH=100>
                <A HREF="catalog-vc/product_list.asp"><IMG SRC="MSCS_Images/admin/products.gif" WIDTH=51 HEIGHT=41 BORDER=0 ALT="Products"></A>
                <BR>
                <A HREF="catalog-vc/product_list.asp"><FONT FACE="Arial, sans-serif">Products</FONT></A>
            </TD>
            <TD ALIGN=CENTER WIDTH=100>
                <A HREF="promo.asp"><IMG SRC="MSCS_Images/admin/promos.gif" WIDTH=51 HEIGHT=41 BORDER=0 ALT="Promos"></A>
                <BR>
                <A HREF="promo.asp"><FONT FACE="Arial, sans-serif">Promos</FONT></A>
            </TD>
        </TR>
        </TABLE>
        </TD>
    <TR>
        <TD ALIGN=LEFT>
            <FONT FACE="Arial, sans-serif" SIZE="+2">T</FONT><FONT FACE="Arial, sans-serif" SIZE="+1">RANSACTIONS</FONT>
        </TD>
        <TD ALIGN=CENTER COLSPAN=3>
        <TABLE>
        <TR>
            <TD ALIGN=CENTER WIDTH=150>
                <A HREF="order-db/order.asp"><IMG SRC="MSCS_Images/admin/orders.gif" WIDTH=51 HEIGHT=41 BORDER=0 ALT="Orders"></A>
                <BR>
                <A HREF="order-db/order.asp"><FONT FACE="Arial, sans-serif">Orders</FONT></A>
            </TD>
            <TD ALIGN=CENTER WIDTH=100>
                <A HREF="shopper-db/shopper.asp"><IMG SRC="MSCS_Images/admin/shoppers.gif" WIDTH=51 HEIGHT=41 BORDER=0 ALT="Shoppers"></A>
                <BR>
                <A HREF="shopper-db/shopper.asp"><FONT FACE="Arial, sans-serif">Shoppers</FONT></A>
            </TD>
        </TR>
        </TABLE>
        </TD>
    </TR>
    <TR>
        <TD ALIGN=LEFT>
            <FONT FACE="Arial, sans-serif" SIZE="+2">S</FONT><FONT FACE="Arial, sans-serif" SIZE="+1">YSTEM</FONT>
        </TD>
        <TD ALIGN=CENTER COLSPAN=3> 
        <% If UBound(PCFfiles) >= LBound(PCFFiles) Then %>
            <FORM METHOD="GET" ACTION="MSCS_PipeEdit/default.asp" TARGET="_top">
                <SELECT NAME="Filename"> 
                    <% For nFile = LBound(PCFfiles,1) To UBound(PCFfiles,1) %>
                        <%= mscsPage.Option(PCFfiles(nFile),0) %> <%= PCFfiles(nFile) %>&nbsp;
                    <% next %>
                </SELECT> 
                <INPUT TYPE="hidden" 
                    NAME="StoreRoot"
                    VALUE="vcturbo">
                <INPUT TYPE="SUBMIT"
                    VALUE=" Edit Pipeline ">
            </FORM>
        <% else %>
            <FONT FACE="Arial, sans-serif" SIZE="-2">No PCF files in <%= ConfigDir %>.<BR>Use PipeEditor.exe to create a new Order Pipeline configuration file.</FONT>
        <% end if %>
        </TD>
    </TR>
    <TR>
        <TD ALIGN=LEFT>
            <FONT FACE="Arial, sans-serif"><B>Store status: [<%= status%>]</B></FONT>
        </TD>
        <TD ALIGN=CENTER>
            <% if status <> "Invalid" then %>
                <FORM METHOD="GET" ACTION="default.asp" TARGET="_top">
                    <INPUT TYPE="SUBMIT"
                           NAME="Status"
                          VALUE="<%= RevStatus %> Store">
                </FORM>
            <% else %>
                *****
            <% end if %>
        </TD>
        <TD ALIGN=CENTER VALIGN=BOTTOM>
            <% if status <> "Invalid" then %>
                <FORM METHOD="GET" ACTION="default.asp" TARGET="_top">
                    <INPUT TYPE="SUBMIT"
                           NAME="Reload"
                          VALUE="Reload Store">
                </FORM>
            <% else %>
                *****
            <% end if %>
        </TD>
        <TD ALIGN=CENTER>
            <A HREF="<% = mscsPage.URL("default.asp") %>" target="_top"><FONT FACE="Arial, sans-serif">Shop</FONT></A>
        </TD>
    </TR>
    <TR>
        <TD ALIGN=LEFT>
            <FONT FACE="Arial, sans-serif"><B>Try our Buy Now:</B></FONT>
        </TD>
        <TD ALIGN=CENTER COLSPAN=3>

            <SCRIPT language="JavaScript">
            <!-- 
            function popUpWindow() {
              BuyNowWin = window.open("../buynow/product.asp?pfid=17","BuyNow","width=400,height=525,scrollbars=yes");
            }
            // -->
            </SCRIPT>
            <A HREF="javascript:popUpWindow()">

            <IMG SRC="../assets/images/buynow.gif" WIDTH=350 ALT="Link to Buy Now control">
             </A>
        </TD>
    </TR>          
</TABLE>

<% REM   footer: %>
<!--#INCLUDE FILE="include/mgmt_footer.asp" -->

<P>
<H6><FONT FACE="Arial, sans-serif">
    The Best Coffee on the Web
    <P>
    The names of companies, products, people, characters, and/or data mentioned herein are fictitious and are in no way intended to represent any real individual, company, product, or event, unless otherwise noted. 
    <P>
    &copy;1996-98 Microsoft Corporation, All rights reserved
</FONT></H6>

</BODY>

</HTML>
