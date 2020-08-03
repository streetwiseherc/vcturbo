<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/shop.asp" -->
<!--#INCLUDE FILE="include/util.asp" -->
<!--#INCLUDE FILE="include/order.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   SHIPPING.ASP                                                            %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<%
REM Run the basic plan
Dim nItemCount
Dim mscsOrderForm 

Set mscsOrderForm = UtilRunBasket(mscsShopperId)
nItemCount = mscsOrderForm.Items.Count
%>

<% ' Include the wallet script code
    Dim strDownlevelURL 
    Dim strPostURL 
    Dim fMSWltPaymentSelector 
    fMSWltAddressSelector = True %>

<!--#INCLUDE FILE="../include/selector.asp" -->

<% strDownlevelURL = pageURL("buynow/shipping.asp") %>
<% strPostURL      = pageSURL("buynow/xt_orderform_prepare.asp") %>

<HTML>
<HEAD>
    <TITLE>Buynow Shipping</TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
    <SCRIPT LANGUAGE="Javascript">
        function submitTheForm()
        {
        <% if fMSWltUplevelBrowser then %>
                if (MSWltPrepareForm(document.shipinfo, 2))
        <% end if %>
        document.shipinfo.submit();
        }
    </SCRIPT>
</HEAD>

<BODY TOPMARGIN=16 BACKGROUND="../assets/images/bnbackground.gif" LANGUAGE="Javascript"
    <% if fMSWltUplevelBrowser and nItemcount > 0 then %>
        LANGUAGE="Javascript"
        ONLOAD="<% = MSWltLoadDone(strDownlevelURL) %>"
    <% end if %>
>

<% ' Header information********************************************************************%>

<H2 ALIGN="left">
    Order Shipping
</H2>

    <% if mscsOrderForm.[_Basket_Errors].Count > 0 then  %>
        <UL> <% Dim errorStr%>
             <% for each errorStr in mscsOrderForm.[_Basket_Errors] %>
                 <li><% = errorStr %></li>
             <% next %>
        </UL>
        <BR>
    <% end if %>
<P>
    <% if CBool(nItemCount = 0) then %>
         <STRONG>Your basket is empty.</STRONG>
    <% else %>
     Please enter the address information below
     and then press the 'Next' button.
</P>
<P>
    <FONT SIZE="+1"><STRONG>Shipping Address</STRONG></FONT>
<P>
    <% if fMSWltUplevelBrowser then %>
        <% ' This table displays the address-picker and contains the form used to post ship info.*****%>
        <TABLE>
            <TR>
                <TD ALIGN=CENTER>
                    <% call MSWltAddressControl(true) %>
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
                </FORM>
                </TD>
            </TR>            
            <TR>
                <TD ALIGN=CENTER>
                <FONT SIZE=1><% = MSWltLastChanceText(strDownlevelURL, "Click here if you have trouble with the Wallet") %></FONT>
                </TD>
            </TR>
        </TABLE>
    <% else %>
        <% ' This table contains the shipping form for downlevel browsers.******************%>
        <TABLE CELLPADDING="0" CELLSPACING="0" BORDER="0" WIDTH=275>
            <FORM NAME="shipinfo" ACTION="<% = strPostURL %>" METHOD="POST">
                <INPUT TYPE=HIDDEN NAME="use_form" VALUE="1">
                <TR>
                    <TD ALIGN=LEFT><B>Name:</B></TD>
                    <TD COLSPAN=3><INPUT TYPE="text" NAME="ship_to_name" SIZE=30,1 MAXLENGTH="50" VALUE=""></TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>Street:</B></TD>
                    <TD COLSPAN=3><INPUT TYPE="text" NAME="ship_to_street" SIZE=30,1 MAXLENGTH="50" VALUE=""></TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>City:</B></TD>
                    <TD COLSPAN=3><INPUT TYPE="text" NAME="ship_to_city" SIZE=30,1 MAXLENGTH="50" VALUE=""></TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>State:</B></TD>
                    <TD><INPUT TYPE="text" NAME="ship_to_state" SIZE=3,1 MAXLENGTH="30" VALUE=""></TD>
                    <TD ALIGN=RIGHT><B>Zip:</B></TD>
                    <TD ALIGN=RIGHT><INPUT TYPE="text" NAME="ship_to_zip" SIZE=10,1 MAXLENGTH="10" VALUE=""></TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>Country:</B></TD>
                    <TD COLSPAN=3><INPUT TYPE="text" NAME="ship_to_country" SIZE=30,1 MAXLENGTH="20" VALUE=""></TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>Phone:</B></TD>
                    <TD COLSPAN=3><INPUT TYPE="text" NAME="ship_to_phone" SIZE=30,1 MAXLENGTH="20" VALUE=""></TD>
                </TR>
                <TR>
                    <TD ALIGN=LEFT><B>E-Mail:</B></TD>
                    <TD COLSPAN=3><INPUT TYPE="text" NAME="ship_to_email" SIZE=30,1 MAXLENGTH="20" VALUE=""></TD>
                </TR>
                <TR><TD>&nbsp;</TD></TR>
    
            </FORM>     
        </TABLE>
     <% end if %>
<% end if %>

<% ' Test for the product in the orderform********************************************* %>
<% if mscsOrderForm.[_Basket_Errors].Count > 0 then  %>
    <UL>
        <% for each errorStr in mscsOrderForm.[_Basket_Errors] %>
            <li><% = errorStr %></li>
        <% next %>
    </UL>
    <P>
<% end if %>

<P>
<CENTER>

<% ' ***** This table displays the product **********************************************%>

<% ' Create Recordset, execute Query to gather the image_filename

   Dim orderitem
   Dim pfid
   Dim rsProduct

   set orderitem = mscsOrderForm.Items(0)
   pfid = CStr(orderitem.pfid)
   set rsProduct = Conn.Execute("select image_filename from vcturbo_product_family where pfid = '" & Replace(pfid, "'", "''") & "'") %>
    
    <TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=300>
        <TR>
            <TD ALIGN="center">
                <IMG SRC="../assets/prodimg/<% = rsProduct("image_filename").Value %>">
            </TD>
            <TD VALIGN="center">
                <STRONG>Description:</STRONG><BR>    
                <FONT size="2">
                    <CENTER>    
                        <% = mscsPage.HTMLEncode(orderitem.quantity) %>&nbsp
                        <% = orderitem.name %>(s) 
                        <% = mscsPage.HTMLEncode(orderitem.attr_text) %>
                    </CENTER>
                </FONT>
                <P>
                <STRONG>Subtotal:</STRONG>
                <FONT size="2"><%  = MSCSDataFunctions.Money(CLng(mscsOrderForm.[_oadjust_subtotal])) %></FONT>
            </TD>
        </TR>
    </TABLE>
<P>
<% ' This table row contains the navigation information************************************%>

<TABLE WIDTH=375>
    <TR>
        <TD ALIGN="left" VALIGN="middle"><P><H3>Step 2 of 4</H3></TD>
        <TD COLSPAN="2" ALIGN="middle" VALIGN="bottom" NOWRAP>
            <FORM METHOD="POST" ACTION="<% = mscsPage.URL("buynow/product.asp","pfid",pfid,"qty",(orderitem.quantity)) %>">
                <INPUT TYPE="SUBMIT" VALUE="&lt;&lt; Previous"> 
                <INPUT TYPE="BUTTON" VALUE="Next &gt;&gt;" onClick="submitTheForm()">&nbsp;&nbsp;
                <INPUT TYPE="BUTTON" VALUE="Cancel"        onClick="self.close()">
            </FORM>
        </TD>    
    </TR>
</TABLE>

</BODY>
</HTML>
