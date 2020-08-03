<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   HEADER.ASP                                                             #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>


<% if IsNull(hd_image) or IsEmpty(hd_image) or hd_image = "" then
    hd_image = "heading_welcome.gif"
end if %>

<% if IsNull(alt) or IsEmpty(alt) or alt  = "" then
    alt = displayName 
end if %>


<TABLE CELLPADDING="2" CELLSPACING="2" WIDTH="600" BORDER="1" RULES="none" background="assets/images/navbar.gif">
    <TR>
        <TD ALIGN="LEFT" WIDTH="115">
            <A HREF="<% = pageURL("home.asp")%>"><IMG SRC="assets/images/logo.gif" WIDTH=98 HEIGHT=76 BORDER=0 ALT="Click here to return Home" ALIGN=TOP></A>
        </TD>
        <% if Not IsNull(mscsShopperId) then %>
            <TD align="center">
               <A HREF="<% = pageURL("listing.asp")%>">      <IMG SRC="assets/images/button_catalog_up.gif"   BORDER=0 ALT="Store Directory" ALIGN=CENTER></A>
               <A HREF="<% = pageURL("search.asp")%>">       <IMG SRC="assets/images/button_search_up.gif"    BORDER=0 ALT="Search" ALIGN=CENTER></A>
               <A HREF="<% = pageURL("basket.asp")%>">       <IMG SRC="assets/images/button_basket_up.gif"    BORDER=0 ALT="Basket" ALIGN=CENTER></A>
               <A HREF="<% = pageURL("checkout-ship.asp")%>"><IMG SRC="assets/images/button_ready2pay_up.gif" BORDER=0 ALT="Pay"    ALIGN=CENTER></A>
           </TD>        
       <% end if %>
    </TR>
</TABLE>

<TABLE WIDTH=600>
    <TR ALIGN=center>
        <TD><BR><IMG SRC="assets/images/<%=hd_image%>" BORDER="0" ALT="<%=alt%>" ALIGN=center><BR></TD>
    </TR>
</TABLE>