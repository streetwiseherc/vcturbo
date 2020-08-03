<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/shop.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   PRODUCT.ASP                                                             %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<HTML>
<HEAD>
    <TITLE><% = displayName %>: Buy Now Wizard</TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">

    <% ' If you support IE2 or Nav1.2 browsers, you'll need to put an if around the script. %>
    
    <SCRIPT LANGUAGE="Javascript">
        function submitTheForm()
        {
            document.prodinfo.submit();
        }
    </SCRIPT>
</HEAD>

<BODY WIDTH="300" BACKGROUND="../assets/images/bnbackground.gif">
<% ' Header information********************************************************************%>
<h2 ALIGN="left">
    Buy now wizard
</h2>

<% REM Create Recordset, execute Query                                                                                                                                                                                
Dim pfid
Dim rsProduct
pfid = Request("pfid")
set rsProduct = Conn.Execute("select pfid, monogramable, list_price, name, image_filename, short_description, long_description from vcturbo_product_family where pfid = '" & Replace(pfid, "'", "''") & "'")
%>

<% if rsProduct.EOF then %>
    Product not available
    <P>
    <FORM>
        <INPUT TYPE=BUTTON VALUE="Close" onClick="self.close()"> 
    </FORM>
<% else %>
    Choose your options and click on next to enter shipping information.
    <P>
    <% ' This table displays the product and options***************************************%>
    <TABLE WIDTH=100%  CELLPADDING=0 CELLSPACING=0 BORDER=0>
        <TR>
            <TD colspan=2><H2><% = mscsPage.HTMLEncode(rsProduct("name").Value) %></H2></TD>
        </TR>
        <TR>
            <TD WIDTH=30% ALIGN=center rowspan=2>
                <% if not isnull(rsProduct("image_filename")) then %>
                    <IMG SRC="../assets/prodimg/<% = rsProduct("image_filename").Value %>">
                <% else %>
                    Image not available
                <% end if %>
            </TD>
            <TD VALIGN=TOP WIDTH=400 colspan=3><% = mscsPage.HTMLEncode(rsProduct("short_description").Value) %></TD>
        </TR>
        <TR>
            <% if not isnull(rsProduct("long_description")) then %>
                <TD VALIGN=TOP WIDTH=400 colspan=3><% = mscsPage.HTMLEncode(rsProduct("long_description").Value) %></TD>
            <% else %>
                <TD VALIGN=TOP WIDTH=400 colspan=3>&nbsp;</TD>
            <% end if %>
        </TR>
        <TR>
            <TD ALIGN=center VALIGN=top><STRONG>Our Price:</STRONG>&nbsp;&nbsp;&nbsp;<FONT SIZE="+2"><% = MSCSDataFunctions.Money(CLng(rsProduct("list_price").Value)) %></FONT></TD>
            <FORM METHOD=POST NAME=prodinfo ACTION="<% = mscsPage.URL("buynow/xt_orderform_additem.asp") %>">
                <INPUT TYPE = HIDDEN NAME="pfid" VALUE="<% = rsProduct("pfid").Value %>">
                <TD VALIGN=top ALIGN=center>
                    <!--#INCLUDE FILE="../include/prd_body.asp" -->
                </TD>
            </FORM>
        </TR>   
        <TR>
            <TD>&nbsp</TD>
        </TR>
        <TR><% ' This table row contains the navigation information************************%>
            <TD ALIGN="left"><H3>Step 1 of 4</H3></TD>
            <TD ALIGN="middle" VALIGN="bottom" colspan="3" nowrap>
                <FORM>
                    <INPUT TYPE=BUTTON   VALUE="Next &gt;&gt;" onClick="submitTheForm()"> 
                    <INPUT TYPE="BUTTON" VALUE="Cancel" onClick="self.close()" >
                </FORM>
            </TD>
        </TR>
    </TABLE>
<% end if %>

</BODY>
</HTML>



