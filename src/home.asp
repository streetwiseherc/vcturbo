<%@ LANGUAGE=vbscript enablesessionstate=false %>
<!--#INCLUDE FILE = "include/shop.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   HOME.ASP                                                               #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>


<HTML>
<HEAD>
<TITLE>Welcome to <% = displayName %></TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>

<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>" BACKGROUND="<% = background %>">

<% hd_image = "heading_history.gif" %>
<% alt = "Home Page" %>
<!--#INCLUDE FILE = "include/header.asp" -->

<TABLE WIDTH="600" BORDER="0">
    <TR>
        <TD colspan=2>
            <P><FONT SIZE=+2>A</FONT>lthough <% = displayName %> has been in business for only six years, it grew from a vision that was 90 years in the making.  Bianca Fiorito, the great granddaughter of Anthony Patricelli, an Italian immigrant to Africa, had grown up watching, listening and learning the ways of planting, picking and brewing coffee beans.  It had been her great grandfather's skill of survival, her grandfather's trade by birth, and her father's vision to one day own a specialty coffee shop of his own.
            <P><FONT SIZE=+2>T</FONT>hose years represent a wealth of knowledge that allowed Bianca to make that dream for her family come true six years ago when she opened up the first <% = displayName %> in Seattle, Washington. Since then, she, along with her father Angelo and grandfather Salvator, has succeeded in opening 15 more stores in the Pacific Northwest. With three generations of collective expertise behind her, she hopes to carry out the traditions of excellence that were instilled in her so long ago.
        </TD>
    </TR>

    <TR>
        <TD>&nbsp;</TD>
    </TR>

    <TR>
        <TD colspan=2>
            <P><FONT SIZE=+1>A</FONT> fixed shipping charge of $9.95 per order will be added to your bill. 
        </TD>
    </TR>
    <TR>
        <TD width="66%">&nbsp;</TD>
        <TD>
            <A HREF = "<% = pageURL("listing.asp")%>"><IMG SRC="assets/images/btn_2_products.gif" BORDER=0 ALT="Products Listing" ALIGN=right></A>
            <A HREF = "<% = pageURL("receipt_list.asp")%>"><IMG SRC="assets/images/btn_2_receipts5.gif" BORDER=0 ALT="Receipts" ALIGN=right></A>
        </TD>
    </TR>
</TABLE>
<% conn.close %>
<!--#INCLUDE FILE = "include/footer.asp" -->

</BODY>
</HTML>
