<%@ LANGUAGE=vbscript enablesessionstate=false %>
<!--#INCLUDE FILE = "include/shop.asp" -->
<!--#INCLUDE FILE = "include/util.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   CHECKOUT-PAY.ASP                                                       #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>

<% Response.ExpiresAbsolute=#Jan 01, 1980 00:00:00# %>

<% 
Dim mscsOrderForm
set mscsOrderForm = UtilRunPlan(mscsShopperId)

if mscsOrderForm.[_Basket_Errors].Count > 0 or mscsOrderForm.[_Purchase_Errors].Count > 0 then %>
    <% 
    dim item
    dim errorList
    set errorList = Server.CreateObject("Commerce.SimpleList")
    for each item in mscsOrderForm.[_Basket_Errors]
        errorList.Add(item)
    next
    for each item in mscsOrderForm.[_Purchase_Errors]
        errorList.Add(item)
    next
%>
    <!--#INCLUDE FILE="include/error.asp" -->
<%  
    Response.End
end if

%>

<% ' Include the wallet script code %>
<% Dim strMSWltAcceptedTypes 
   Dim strTotal 
   Dim fMSWltPaymentSelector 
   Dim strDownlevelURL 
   Dim strPostURL %>



<% strMSWltAcceptedTypes = "visa:clear;mastercard:clear" %>
<% strTotal = MSCSDataFunctions.Money(CLng(mscsOrderForm.[_total_total])) %>
<% fMSWltPaymentSelector = True %>

<% strDownlevelURL = pageURL("checkout-pay.asp") %>
<% strPostURL      = pageSURL("xt_orderform_prepare.asp") %>


<HTML>

<HEAD>
    <!--#INCLUDE FILE = "include/selector.asp" -->
    
    <SCRIPT LANGUAGE="Javascript">
        function submitPayInfo()
        {
            if (MSWltPrepareForm(document.payinfo, 2))
                document.payinfo.submit();
        }
    </SCRIPT>
        
    <TITLE>Checkout - Pay</TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>

<% hd_image = "heading_checkout.gif" %>
<% alt = "Checkout - Pay" %>


<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>" BACKGROUND="<% = background %>"
    <% if fMSWltUplevelBrowser then %>
        LANGUAGE="Javascript"
        ONLOAD="<% = MSWltLoadDone(pageSURL("checkout-pay.asp")) %>"
    <% end if %>
>
<!--#INCLUDE FILE = "include/header.asp" -->

<%' Text Table%>
<TABLE WIDTH=600>
    <TR>
        <TD>
        Your purchase will cost <% = strTotal %>. Please enter your
        credit card information and press the Purchase Now button below to 
        complete your purchase.        
        <P>
        If you want to change your order before proceeding, go to the
        <A HREF = "<% = pageURL("basket.asp") %>">shopping basket</A>.
        <P>
        </TD>
    </TR>
</TABLE>

<TABLE WIDTH=350>
    <TR>
        <TD>&nbsp;</TD>
        <TD COLSPAN=2 WIDTH=140><FONT SIZE=+1><STRONG>Total Charges:</STRONG></FONT></TD>
        <TD WIDTH=50>&nbsp;</TD>
    </TR>
    <TR>
        <TD>&nbsp;</TD>
        <TD ALIGN=LEFT WIDTH=80><B>Subtotal:</B></TD>
        <TD ALIGN=RIGHT>
            <% =MSCSDataFunctions.Money(CLng(mscsOrderForm.[_oadjust_subtotal])) %>
        </TD>
    </TR>
    <TR>
        <TD>&nbsp;</TD>
        <TD ALIGN=LEFT><B>Tax:</B></TD>
        <TD ALIGN=RIGHT>
            <% =MSCSDataFunctions.Money(CLng(mscsOrderForm.[_tax_total])) %>
        </TD>
    </TR>
    <TR>
        <TD>&nbsp;</TD>
        <TD ALIGN=LEFT><B>Shipping:</B></TD>
        <TD ALIGN=RIGHT>
            <% =MSCSDataFunctions.Money(CLng(mscsOrderForm.[_shipping_total])) %>
        </TD>
    </TR>
    <TR>
        <TD>&nbsp;</TD>
        <TD COLSPAN=2><HR></TD>
    </TR>
    <TR>
        <TD>&nbsp;</TD>
        <TD ALIGN=LEFT><B>TOTAL:</B></TD>
        <TD ALIGN=RIGHT>
            <B><% =MSCSDataFunctions.Money(CLng(mscsOrderForm.[_total_total])) %></B>
        </TD>
    </TR>
</TABLE>

<P>
<FONT SIZE="+1"><STRONG>Payment Information</STRONG></FONT>
<P>
<%
if fMSWltUplevelBrowser then %>
            <TABLE BORDER=0 CELLSPACING=2 CELLPADDING=2>
                <TR><TD ALIGN=CENTER>
                    <% call MSWltPaymentControl(False, strMSWltAcceptedTypes, strTotal) %>
                </TD></TR>
                <TR>
                    <TD ALIGN=CENTER>
                        <FORM NAME="payinfo" ACTION="<% = pageSURL("xt_orderform_purchase.asp") %>" METHOD=POST>
                            <% = mscsPage.Verifywith(mscsOrderForm,"_total_total","ship_to_zip","_tax_total") %>
                            <INPUT TYPE=HIDDEN NAME="bill_to_name">
                            <INPUT TYPE=HIDDEN NAME="bill_to_street">
                            <INPUT TYPE=HIDDEN NAME="bill_to_city">
                            <INPUT TYPE=HIDDEN NAME="bill_to_state">
                            <INPUT TYPE=HIDDEN NAME="bill_to_zip">
                            <INPUT TYPE=HIDDEN NAME="bill_to_country">
                            <INPUT TYPE=HIDDEN NAME="bill_to_phone">
                            <INPUT TYPE=HIDDEN NAME="bill_to_email">
                            <INPUT TYPE=HIDDEN NAME="cc_name">
                            <INPUT TYPE=HIDDEN NAME="cc_type">
                            <INPUT TYPE=HIDDEN NAME="_cc_number">
                            <INPUT TYPE=HIDDEN NAME="_cc_expmonth">
                            <INPUT TYPE=HIDDEN NAME="_cc_expyear">
                            <INPUT TYPE=BUTTON VALUE="Purchase Now" onClick="submitPayInfo()">
                        </FORM>
                    </TD>
                </TR>
                <TR>
                    <TD ALIGN=CENTER>
                        <FONT SIZE=1>
                           <% = MSWltLastChanceText(strDownlevelURL, "Click here if you have trouble with the Wallet") %>
                        </FONT>
                    </TD>
                </TR>
            </TABLE>
<% else %>
    <FORM NAME="payinfo" ACTION="<% = pageSURL("xt_orderform_purchase.asp") %>" METHOD=POST>
        <% = mscsPage.Verifywith(mscsOrderForm,"_total_total","ship_to_zip","_tax_total") %>
        <TABLE BORDER=0>
               <TR>
                    <TD ALIGN=LEFT COLSPAN=5>
                        <FONT SIZE="+1"><STRONG>Card Type</STRONG></FONT>
                    </TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>Name On Card:</B></TD>
                    <TD COLSPAN=4>
                        <INPUT TYPE="text" NAME="cc_name" MAXLENGTH="100" VALUE="<% = mscsPage.HTMLEncode(mscsOrderForm.cc_name) %>" SIZE=45,1>
                    </TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>Card Number:</B></TD>
                    <TD COLSPAN=4>
                        <INPUT TYPE="text" NAME="_cc_number" MAXLENGTH="19" SIZE=45,1>
                    </TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>Type:</B></TD>
                    <TD>
                        <SELECT NAME="cc_type">
                            <% = mscsPage.Option("Visa","cc_type")%>Visa
                            <% = mscsPage.Option("Mastercard","cc_type")%>MasterCard&nbsp;&nbsp;
                        </SELECT>
                    </TD>        
                        <TD ALIGN=RIGHT><B>Expiration:</B></TD>
                    <TD><% Dim iYear%>
                        <% Dim iMonth%>
                        <% iYear = Year(now) %>
                        <% iMonth = Month(now) %>
                        <SELECT NAME="_cc_expmonth">
                            <% = mscsPage.Option( 1 , iMonth )%>Jan
                            <% = mscsPage.Option( 2 , iMonth )%>Feb
                            <% = mscsPage.Option( 3 , iMonth )%>Mar
                            <% = mscsPage.Option( 4 , iMonth )%>Apr
                            <% = mscsPage.Option( 5 , iMonth )%>May
                            <% = mscsPage.Option( 6 , iMonth )%>Jun
                            <% = mscsPage.Option( 7 , iMonth )%>Jul
                            <% = mscsPage.Option( 8 , iMonth )%>Aug
                            <% = mscsPage.Option( 9 , iMonth )%>Sep
                            <% = mscsPage.Option(10 , iMonth )%>Oct
                            <% = mscsPage.Option(11 , iMonth )%>Nov
                            <% = mscsPage.Option(12 , iMonth )%>Dec
                        </SELECT>
                    </TD>
                    <TD>
                        <SELECT NAME="_cc_expyear">
                            <% = mscsPage.Option(1997 , iYear ) %>1997
                            <% = mscsPage.Option(1998 , iYear ) %>1998
                            <% = mscsPage.Option(1999 , iYear ) %>1999
                            <% = mscsPage.Option(2000 , iYear ) %>2000
                            <% = mscsPage.Option(2001 , iYear ) %>2001
                            <% = mscsPage.Option(2002 , iYear ) %>2002
                            <% = mscsPage.Option(2003 , iYear ) %>2003
                        </SELECT>
                    </TD>
                <TR>
                    <TD COLSPAN=5>&nbsp;</TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT COLSPAN=5>
                        <FONT SIZE="+1"><STRONG>Billing Address</STRONG></FONT>
                    </TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>Name:</B></TD>
                    <TD COLSPAN=4>
                        <INPUT TYPE="text" NAME="bill_to_name" SIZE=45,1 MAXLENGTH="50" VALUE="<% = mscsOrderForm.bill_to_name %>"> 
                    </TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>Street:</B></TD>
                    <TD COLSPAN=4>
                        <INPUT TYPE="text" NAME="bill_to_street" SIZE=45,1 MAXLENGTH="50" VALUE="<% = mscsOrderForm.bill_to_street %>">
                    </TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>City:</B></TD>
                    <TD COLSPAN=4>
                        <INPUT TYPE="text" NAME="bill_to_city" SIZE=45,1 MAXLENGTH="50" VALUE="<% = mscsOrderForm.bill_to_city %>">
                    </TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>State:</B></TD>
                    <TD COLSPAN=2>
                        <INPUT TYPE="text" NAME="bill_to_state" SIZE=23,1 MAXLENGTH="30" VALUE="<% = mscsOrderForm.bill_to_state%>">
                    </TD>
                    <TD ALIGN=LEFT><B>Zip:</B></TD>
                    <TD COLSPAN=0 ALIGN=RIGHT>
                        <INPUT TYPE="text" NAME="bill_to_zip" SIZE=10,1 MAXLENGTH="10" VALUE="<% = mscsOrderForm.bill_to_zip%>">
                    </TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>Country:</B></TD>
                    <TD COLSPAN=4>
                        <INPUT TYPE="text" NAME="bill_to_country" SIZE=45,1 MAXLENGTH="20" VALUE="<% = mscsOrderForm.bill_to_country %>">
                    </TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>Phone:</B></TD>
                    <TD COLSPAN=4>
                        <INPUT TYPE="text" NAME="bill_to_phone" SIZE=45,1 MAXLENGTH="20" VALUE="<% = mscsOrderForm.bill_to_phone%>">
                    </TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>Email:</B></TD>
                    <TD COLSPAN=4>
                        <INPUT TYPE="text" NAME="bill_to_email" SIZE=45,1 MAXLENGTH="20" VALUE="<% = mscsOrderForm.bill_to_email %>">
                    </TD>
                </TR>
                <TR><TD COLSPAN=5>&nbsp;</TD></TR>               
                <TR>
                    <TD>&nbsp;</TD>
                    <TD ALIGN=RIGHT COLSPAN=4 WIDTH=310>
                        <INPUT TYPE=RESET VALUE="Reset Form">  
                        <INPUT TYPE=SUBMIT VALUE="Purchase">
                    </TD>
                </TR>
        </TABLE>
    </FORM>
<% end if %>
<% conn.close %>            
<!--#INCLUDE FILE = "include/footer.asp" -->

</BODY>
</HTML>
       
