<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/shop.asp" -->
<!--#INCLUDE FILE="include/util.asp" -->
<!--#INCLUDE FILE="include/order.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   PAYMENT.ASP                                                             %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% 
Dim mscsOrderForm
set mscsOrderForm = UtilRunPlan(mscsShopperId)

if mscsOrderForm.[_Basket_Errors].Count > 0 or mscsOrderForm.[_Purchase_Errors].Count > 0 then
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
<% REM strTotal = MSCSDataFunctions.Money(CLng(mscsOrderForm.[_total_total])) %>
<% fMSWltPaymentSelector = True %>

<% strDownlevelURL = pageURL("buynow/payment.asp") %>
<% strPostURL      = pageSURL("buynow/xt_orderform_purchase.asp") %>

<HTML>
<HEAD>
    <!--#INCLUDE FILE="../include/selector.asp" -->
    <TITLE><% = displayName %>: Final Purchase Summary</TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
    <% ' If you support IE2 or Nav1.2 browsers, you'll need to put an if around the script. %>
    <SCRIPT LANGUAGE="Javascript">
        function submitTheForm()
        {
            <% if fMSWltUplevelBrowser then %>
                if (MSWltPrepareForm(document.payinfo, 2))
            <% end if %>
                    document.payinfo.submit()
        }
    </SCRIPT>
</HEAD>

<BODY BACKGROUND="../assets/images/bnbackground.gif" LANGUAGE="Javascript"
    <% if fMSWltUplevelBrowser then %>
        LANGUAGE="Javascript"
        ONLOAD="<% = MSWltLoadDone(pageSURL("buynow/payment.asp")) %>"
    <% end if %>
>
<% ' Header information********************************************************************%>
<H2>
    <% if fMSWltUplevelBrowser then %>
        Select Payment
    <% else %>
        Order Payment
    <% end if %>
</H2>
<P>
    <% if fMSWltUplevelBrowser then %>
        Choose how to pay from the payment methods drop down then click next to complete the purchase.
    <% else %>
        Enter your billing address and credit card information then click next to complete the purchase.
    <% end if %>
</P>
<P>
    <FONT SIZE="+1"><STRONG>Payment Information</STRONG></FONT>
<P>
<%
if fMSWltUplevelBrowser then %>
    <% ' This table displays the wallet and contains the form used to post billing info.*****%>
    <TABLE BORDER=0>
        <TR>
            <TD ALIGN=CENTER>
                <% call MSWltPaymentControl(true, strMSWltAcceptedTypes, strTotal) %>
            <FORM NAME="payinfo" ACTION="<% = pageSURL("buynow/xt_orderform_purchase.asp") %>" METHOD=POST>
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
    <% ' This table contains the billing form for downlevel browsers.******************%>
    <P>
    <STRONG>
        Billing Address:
    </STRONG>
    <BR>
    <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="2">
        <FORM NAME="payinfo" ACTION="<% = pageSURL("buynow/xt_orderform_purchase.asp") %>" METHOD=POST>
        <% = mscsPage.Verifywith(mscsOrderForm,"_total_total","ship_to_zip","_tax_total") %>
        <INPUT TYPE="HIDDEN" NAME="bill_to_country" VALUE="USA">
        <TR>
            <TD ALIGN=RIGHT>Name:</TD>
            <TD COLSPAN=3><INPUT TYPE="text" NAME="bill_to_name" VALUE="<% = mscsPage.HTMLEncode(mscsOrderForm.ship_to_name) %>" SIZE=30,1></TD>
        </TR>
        <TR>
            <TD ALIGN=RIGHT>Street:</TD>
            <TD COLSPAN=3><INPUT TYPE="text" NAME="bill_to_street" VALUE="<% = mscsPage.HTMLEncode(mscsOrderForm.ship_to_street) %>" SIZE=30,1></TD>
        </TR>
        <TR>
            <TD ALIGN=RIGHT>City:</TD>
            <TD COLSPAN=3><INPUT TYPE="text" NAME="bill_to_city" VALUE="<% = mscsPage.HTMLEncode(mscsOrderForm.ship_to_city) %>" SIZE=30,1></TD>
        </TR>
        <TR>
            <TD ALIGN=RIGHT>State:</TD>
            <TD><INPUT TYPE="text" NAME="bill_to_state" VALUE="<% = mscsPage.HTMLEncode(mscsOrderForm.ship_to_state) %>" SIZE=3,1>&nbsp;</TD>
            <TD ALIGN=RIGHT>ZIP:<INPUT TYPE="text" NAME="bill_to_zip" VALUE="<% = mscsPage.HTMLEncode(mscsOrderForm.ship_to_zip) %>" SIZE=10,1>&nbsp;</TD>
        </TR>
        <TR>
            <TD ALIGN=RIGHT>Phone:</TD>
            <TD COLSPAN=3><INPUT TYPE="text" NAME="bill_to_phone" VALUE="<% = mscsPage.HTMLEncode(mscsOrderForm.ship_to_phone) %>" SIZE=30,1></TD>
        </TR>
        <TR>
            <TD ALIGN=RIGHT>Email</TD>
            <TD COLSPAN=3><INPUT TYPE="text" NAME="bill_to_email" VALUE="<% = mscsPage.HTMLEncode(mscsOrderForm.ship_to_email) %>" SIZE=30,1></TD>
        </TR>
    </TABLE>
    <BR>
    <P>
    <STRONG>
        Credit Card Info:
    </STRONG>
    <BR>
    <TABLE BORDER="0" CELLSPACING="0" CELLPADDING="2" WIDTH=350>
        <TR>
            <TD ALIGN=RIGHT>Name on card:</TD>
            <TD COLSPAN=4><INPUT TYPE="text" NAME="cc_name" VALUE="<% = mscsPage.HTMLEncode(mscsOrderForm.bill_to_name) %>" SIZE=30,1></TD>
        </TR>
        <TR>
            <TD ALIGN=RIGHT>Card #:</TD>
            <TD COLSPAN=4><INPUT TYPE="text" NAME="_cc_number" SIZE=30,1></TD>
        </TR>
        <TR>
            <TD ALIGN=RIGHT>Card type:</TD>
            <TD>
            <SELECT NAME="cc_type">
                <OPTION VALUE="Visa"> VISA
                <OPTION VALUE="Mastercard"> MasterCard
                <OPTION VALUE="AMEX"> American Express
            </SELECT>
            </TD>
            <TD ALIGN=RIGHT>Exp. date:</TD>
                    <% Dim iYear
                       Dim iMonth

                       iYear = Year(Date) 
                       iMonth = Month(Date) %>
            <TD><SELECT NAME="_cc_expmonth">
                    <% = mscsPage.Option(1, iMonth) %> Jan
                    <% = mscsPage.Option(2, iMonth) %> Feb
                    <% = mscsPage.Option(3, iMonth) %> Mar
                    <% = mscsPage.Option(4, iMonth) %> Apr
                    <% = mscsPage.Option(5, iMonth) %> May
                    <% = mscsPage.Option(6, iMonth) %> Jun
                    <% = mscsPage.Option(7, iMonth) %> Jul
                    <% = mscsPage.Option(8, iMonth) %> Aug
                    <% = mscsPage.Option(9, iMonth) %> Sep
                    <% = mscsPage.Option(10, iMonth) %> Oct
                    <% = mscsPage.Option(11, iMonth) %> Nov
                    <% = mscsPage.Option(12, iMonth) %> Dec
                </SELECT>
                <SELECT NAME="_cc_expyear">
                    <% = mscsPage.Option(1996, iYear) %> 1996
                    <% = mscsPage.Option(1997, iYear) %> 1997
                    <% = mscsPage.Option(1998, iYear) %> 1998
                    <% = mscsPage.Option(1999, iYear) %> 1999
                    <% = mscsPage.Option(2000, iYear) %> 2000
                    <% = mscsPage.Option(2001, iYear) %> 2001
                    <% = mscsPage.Option(2002, iYear) %> 2002
                    <% = mscsPage.Option(2003, iYear) %> 2003
                </SELECT>
            </TD>
        </TR>
    </TABLE>
    <BR>
    </FORM>

<% end if %>
<P>
<% Dim orderitem %>
<% set orderitem = mscsOrderForm.Items(0) 
    Dim pfid
    Dim rsProduct
    pfid = CStr(orderitem.pfid)  %>
    <TABLE BORDER=0 CELLSPACING="0" CELLPADDING="0" WIDTH="350">
        <TR>
            <TD>
            <% REM Create Recordset, execute Query to gather the image_filename
            set rsProduct = Conn.Execute("select image_filename from vcturbo_product_family where pfid = '" & Replace(pfid, "'", "''") & "'") %>
            <IMG SRC="../assets/prodimg/<% = rsProduct("image_filename").Value %>">
            </TD>
            <TD VALIGN="center">
                <STRONG>
                    Description:
                </STRONG>
                <BR>
                <FONT SIZE="2">
                    <% = mscsPage.HTMLEncode(orderitem.quantity) %>&nbsp
                    <% = orderitem.name %>(s) 
                    <% = mscsPage.HTMLEncode(orderitem.attr_text) %>
                </FONT><P> 
                <STRONG>
                    Shipping To:
                </STRONG>
                <BR>
                <% = mscsPage.HTMLEncode(mscsOrderForm.ship_to_name) %><BR> 
                <% = mscsPage.HTMLEncode(mscsOrderForm.ship_to_street) %><BR>
                <% = mscsPage.HTMLEncode(mscsOrderForm.ship_to_city) %>, 
                <% = mscsPage.HTMLEncode(mscsOrderForm.ship_to_state) & " " %>
                <% = mscsPage.HTMLEncode(mscsOrderForm.ship_to_zip) %>
                <BR>
                <CENTER>
                <TD ALIGN=right>
                <TABLE BORDER="1" CELLSPACING="0" CELLPADDING="3" WIDTH="75">
                    <TR>
                        <TD VALIGN="top"><p ALIGN="LEFT">
                            Subtotal:
                        </TD>
                        <TD ALIGN="right" VALIGN="top" nowrap>
                            <P ALIGN="right">
                                <% = MSCSDataFunctions.Money(CLng(mscsOrderForm.[_oadjust_subtotal])) %>
                        </TD>
                    </TR>
                    <TR>
                        <TD VALIGN="top">
                            <P ALIGN="LEFT">
                                 Shipping:
                        </TD>
                        <TD ALIGN="right" VALIGN="top">
                            <% = MSCSDataFunctions.Money(CLng(mscsOrderForm.[_shipping_total])) %>
                        </TD>
                    </TR>
                    <TR>
                        <TD VALIGN="top"><p ALIGN="LEFT">
                            Tax:
                        </TD>
                        <TD ALIGN="right" VALIGN="top">
                            <% = MSCSDataFunctions.Money(CLng(mscsOrderForm.[_tax_total])) %>
                        </TD>
                    </TR>
                    <TR>
                        <TD VALIGN="top"><p ALIGN="LEFT">
                            Total:
                        </TD>
                        <TD ALIGN="right" VALIGN="top">
                            <% = MSCSDataFunctions.Money(CLng(mscsOrderForm.[_total_total])) %>
                        </TD>
                    </TR>
                </TABLE>

            </TR>
        </TABLE>    


            <% ' Put button below everything else. %>
    </TABLE>    
    <TABLE WIDTH=375>    
        <TR VALIGN="bottom">
            <TD ALIGN="left" VALIGN="middle"><P><H3>Step 3 of 4</H3></TD>
            <TD colspan="2" ALIGN="middle" VALIGN="bottom" nowrap>
                <FORM METHOD="POST" ACTION="<% = mscsPage.SURL("buynow/shipping.asp", "pfid", pfid, "use_form", 0) %>">
                    <INPUT TYPE=SUBMIT   VALUE="&lt;&lt; Previous" > 
                    <INPUT TYPE=BUTTON   VALUE="Next &gt;&gt;" onClick="submitTheForm()">&nbsp;&nbsp;
                    <INPUT TYPE="BUTTON" VALUE="Cancel" onClick="self.close()" >
                </FORM>
            </TD>
        </TR>
    </TABLE>


</BODY>
</HTML>
