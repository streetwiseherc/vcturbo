<%@ LANGUAGE=vbscript enablesessionstate=false %>
<!--#INCLUDE FILE = "include/shop.asp" -->
<!--#INCLUDE FILE = "include/util.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   CHECKOUT-SHIP.ASP                                                      #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>
<% Response.ExpiresAbsolute=#Jan 01, 1980 00:00:00# %>

<%
REM Run the basic plan
Dim mscsOrderForm
Dim nItemCount

Set mscsOrderForm = UtilRunPlan(mscsShopperId)
nItemCount = mscsOrderForm.Items.Count
%>

<% ' Include the wallet script code

    Dim strDownlevelURL 
    Dim strPostURL 
    Dim fMSWltPaymentSelector 



    fMSWltAddressSelector = True %>
<!--#INCLUDE FILE="include/selector.asp" -->

<% strDownlevelURL = pageURL("checkout-ship.asp") %>
<% strPostURL      = pageSURL("xt_orderform_prepare.asp") %>


<HTML>

<HEAD>
    <SCRIPT LANGUAGE="Javascript">
    function submitShipToAddr()
    {
        if (MSWltPrepareForm(document.shipinfo, 2))
            document.shipinfo.submit();
    }
    </SCRIPT>    

    <TITLE>Checkout - Shipping</TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>
<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>" BACKGROUND="<% = background %>"
    <% if fMSWltUplevelBrowser and nItemcount > 0 then %>
        LANGUAGE="Javascript"
        ONLOAD="<% = MSWltLoadDone(strDownlevelURL) %>"
    <% end if %>
>

<% hd_image = "heading_checkout.gif" %>
<% alt = "Checkout - Shipping" %>
<!--#INCLUDE FILE = "include/header.asp" -->

<% if mscsOrderForm.[_Basket_Errors].Count > 0 then  %>
    <UL> <% Dim errorStr%>
         <% for each errorStr in mscsOrderForm.[_Basket_Errors] %>
             <li><% = errorStr %></li>
         <% next %>
    </UL>
    <BR>
<% end if %>
<% if CBool(nItemCount = 0) then %>
     <STRONG>Your basket is empty.</STRONG>
<% else %>
     Please enter the address information below
     and then press the 'Continue' button.
     <P>
     If you want to change your order before proceeding, go to the
     <A HREF = "<% = pageURL("basket.asp") %>">shopping basket</A>.
     <P>
     <FONT SIZE="+1"><STRONG>Shipping Address</STRONG></FONT>
     <P>
     <% 
     if fMSWltUplevelBrowser then %>
        <TABLE>
            <TR><TD ALIGN=CENTER>
               <% Dim small%>
               <% call MSWltAddressControl(small) %>
            </TD></TR>
            <TR><TD ALIGN=CENTER>
                <FORM NAME="shipinfo" ACTION="<% = strPostURL %>" METHOD=POST>
                    <INPUT TYPE=HIDDEN NAME="use_form"        VALUE = "0" >
                    <INPUT TYPE=HIDDEN NAME="ship_to_country" VALUE = "USA">
                    <INPUT TYPE=HIDDEN NAME="ship_to_name">
                    <INPUT TYPE=HIDDEN NAME="ship_to_street">
                    <INPUT TYPE=HIDDEN NAME="ship_to_city">
                    <INPUT TYPE=HIDDEN NAME="ship_to_state">
                    <INPUT TYPE=HIDDEN NAME="ship_to_zip">
                    <INPUT TYPE=HIDDEN NAME="ship_to_phone">
                    <INPUT TYPE=HIDDEN NAME="ship_to_email">
                    <INPUT TYPE=BUTTON VALUE="Continue" onClick="submitShipToAddr()">
                </FORM>
            </TD></TR>
            <TR>
                <TD ALIGN=CENTER>
                <FONT SIZE=1><% = MSWltLastChanceText(strDownlevelURL, "Click here if you have trouble with the Wallet") %></FONT>
                </TD>
            </TR>
        </TABLE>
    <% else %>
        <TABLE CELLPADDING="0" CELLSPACING="0" BORDER="0" WIDTH=500>
            <FORM NAME="shipinfo" ACTION="<% = strPostURL %>" METHOD="POST">
                <INPUT TYPE=HIDDEN NAME="use_form" VALUE="1">
                <TR>
                    <TD ALIGN=LEFT><B>Name:</B></TD>
                    <TD COLSPAN=3>
                        <INPUT TYPE="text" NAME="ship_to_name" SIZE=54,1 MAXLENGTH="50" VALUE="<% = mscsOrderForm.ship_to_name %>">
                    </TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>Street:</B></TD>
                    <TD COLSPAN=3>
                        <INPUT TYPE="text" NAME="ship_to_street" SIZE=54,1 MAXLENGTH="50" VALUE="<% = mscsOrderForm.ship_to_street %>">
                    </TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>City:</B></TD>
                    <TD COLSPAN=3>
                        <INPUT TYPE="text" NAME="ship_to_city" SIZE=54,1 MAXLENGTH="50" VALUE="<% = mscsOrderForm.ship_to_city %>">
                    </TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>State:</B></TD>
                    <TD>
                        <INPUT TYPE="text" NAME="ship_to_state" SIZE=23,1 MAXLENGTH="30" VALUE="<% = mscsOrderForm.ship_to_state %>">
                    </TD>
                    <TD ALIGN=RIGHT><B>Zip:</B></TD>
                    <TD>
                        <INPUT TYPE="text" NAME="ship_to_zip" SIZE=10,1 MAXLENGTH="10" VALUE="<% = mscsOrderForm.ship_to_zip %>">
                    </TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>Country:</B></TD>
                    <TD COLSPAN=3>
                        <INPUT TYPE="text" NAME="ship_to_country" SIZE=54,1 MAXLENGTH="20" VALUE="<% = mscsOrderForm.ship_to_country %>">
                    </TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>Phone:</B></TD>
                    <TD COLSPAN=3>
                        <INPUT TYPE="text" NAME="ship_to_phone" SIZE=54,1 MAXLENGTH="20" VALUE="<% = mscsOrderForm.ship_to_phone %>">
                    </TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>E-Mail:</B></TD>
                    <TD COLSPAN=3>
                        <INPUT TYPE="text" NAME="ship_to_email" SIZE=54,1 MAXLENGTH="20" VALUE="<% = mscsOrderForm.ship_to_email %>">
                    </TD>
                </TR>
                <TR><TD>&nbsp;</TD></TR>
                <TR>
                    <TD>&nbsp;</TD>
                    <TD ALIGN=RIGHT COLSPAN=3 WIDTH=350>
                        <INPUT TYPE="reset" VALUE="Reset Form">
                        <INPUT TYPE="SUBMIT" VALUE="Total/Finish">
                    </TD>
                </TR>       
            </FORM>     
        </TABLE>
     <% end if %>
<% end if %>
<% conn.close %>
<!--#INCLUDE FILE = "include/footer.asp" -->

</BODY>
</HTML>
