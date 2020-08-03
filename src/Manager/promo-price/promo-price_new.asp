<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   PROMO-PRICE_NEW.ASP                                                     %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% REM This code fills in the data to be used for the shopper and product selection boxes


    REM This list is the fields that are available for product selection
    set productList = Server.CreateObject("Commerce.SimpleList")
    productList.Add("sku")
    productList.Add("name")
    productList.Add("list_price")
    productList.Add("pfid")
    
    REM This list is the fields that are available for shopper selection
    set shopperList = Server.CreateObject("Commerce.SimpleList")
    cmdTemp.CommandText = "select * from vcturbo_shopper where 1 = 2"
    Set rsShoppercol = Server.CreateObject("ADODB.Recordset")
    rsShoppercol.Open cmdTemp, , adOpenKeyset, adLockReadOnly
    on error resume next
    for i = 0 to rsShoppercol.Fields.Count
        shopperList.Add(rsShoppercol(i).Name)
    next
    on error goto 0
    rsShopperCol.Close

%>

<%
    cmdTemp.CommandText = "SELECT pfid, name FROM vcturbo_product_family ORDER BY name"
    Set rsProduct = Server.CreateObject("ADODB.Recordset")
    rsProduct.Open cmdTemp, , adOpenKeyset, adLockReadOnly

    sub productOptions()
        rsProduct.Movefirst
        While Not rsProduct.EOF
            Response.write mscsPage.Option(Cstr(rsProduct("pfid")), 0) & " " & rsProduct("name") & "&nbsp;" & vbCr
            rsProduct.MoveNext
        Wend
    end sub

    PromoType = mscsPage.RequestString("promo_type","100")
%>

<% pageTitle = "New Price Promotion" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<TABLE BORDER="1" CELLPADDING="5" BORDERCOLOR="#000000">

<% Select Case PromoType '===============================================
    REM -- buy x get y at z% off-----------------------------------------------
    Case "1" %>
        <SCRIPT LANGUAGE="JavaScript">
        <!--
            function promoform_onsubmit() {
                choices = new Array(5);
                choices[0] = document.promoform.cond_value.options[document.promoform.cond_value.selectedIndex].value;
                choices[1] = document.promoform.award_value.options[document.promoform.award_value.selectedIndex].value;
                choices[2] = document.promoform.disc_value.options[document.promoform.disc_value.selectedIndex].value;
                choices[3] = document.promoform.cond_min.value;
                choices[4] = document.promoform.award_max.value;
                description = "Buy %3 '%0' get  %4 '%1' at %2 % OFF!";
                for (var i = 0; i < choices.length; i++) {
                    description = description.substring(0,description.indexOf("%" + i)) + choices[i] + description.substring(description.indexOf("%" + i) + 2,description.length);
                }
            
                document.promoform.promo_description.value = description;
                return true;
            }
        //-->
        </SCRIPT>
    <% REM -- buy x get y at $z off---------------------------------------------
    Case "2" %>
        <SCRIPT LANGUAGE="JavaScript">
        <!--
            function promoform_onsubmit() {
                choices = new Array(5);
                choices[0] = document.promoform.cond_value.options[document.promoform.cond_value.selectedIndex].value;
                choices[1] = document.promoform.award_value.options[document.promoform.award_value.selectedIndex].value;
                choices[2] = document.promoform.disc_value.options[document.promoform.disc_value.selectedIndex].value;
                choices[3] = document.promoform.cond_min.value;
                choices[4] = document.promoform.award_max.value;
                description = "Buy %3 '%0' get %4 '%1' at %2 $ OFF!";
                for (var i = 0; i < choices.length; i++) {
                    description = description.substring(0,description.indexOf("%" + i)) + choices[i] + description.substring(description.indexOf("%" + i) + 2,description.length);
                }
            
                document.promoform.promo_description.value = description;

                return true;
            }
        //-->
        </SCRIPT>
    <% REM -- buy 2 x for the price of 1---------------------------------------
    Case "3" %>
        <SCRIPT LANGUAGE="JavaScript">
        <!--
            function promoform_onsubmit() {
                x = document.promoform.cond_value.options[document.promoform.cond_value.selectedIndex].value;
                description = "Buy 2 '%x' for the price of 1!";
                document.promoform.promo_description.value = description.substring(0,description.indexOf("%x")) + x + description.substring(description.indexOf("%x") + 2,description.length)

                document.promoform.award_value.value = x;
                return true;
            }
        //-->
        </SCRIPT>
<% End Select '================================================================ %>
    <FORM METHOD="POST" NAME="promoform"
        <% if PromoType <> "100" then %> onSubmit="promoform_onsubmit()"<% end if %>
        ACTION="xt_data_add_update.asp">

<% Select Case PromoType '===============================================
    REM -- buy x get y at z% off-----------------------------------------------
    Case "1" %>
        <INPUT TYPE="HIDDEN" NAME="promo_type"   VALUE="1">
        <INPUT TYPE="HIDDEN" NAME="promo_description">
        <INPUT TYPE="HIDDEN" NAME="shopper_all"  VALUE="1">
        <INPUT TYPE="HIDDEN" NAME="cond_basis"   VALUE="Q">
        <INPUT TYPE="HIDDEN" NAME="cond_all"     VALUE="0">
        <INPUT TYPE="HIDDEN" NAME="cond_op"      VALUE="=">
        <INPUT TYPE="HIDDEN" NAME="cond_column"  VALUE="pfid">
        <INPUT TYPE="HIDDEN" NAME="award_op"     VALUE="=">
        <INPUT TYPE="HIDDEN" NAME="award_all"    VALUE="0">
        <INPUT TYPE="HIDDEN" NAME="award_column" VALUE="pfid">
    <INPUT TYPE="HIDDEN" NAME="disjoint_cond_award" VALUE="0">
        <INPUT TYPE="HIDDEN" NAME="disc_type"    VALUE="%">
    <% REM -- buy x get y at $z off---------------------------------------------
    Case "2" %>
        <INPUT TYPE="HIDDEN" NAME="promo_type"   VALUE="2">
        <INPUT TYPE="HIDDEN" NAME="promo_description">
        <INPUT TYPE="HIDDEN" NAME="shopper_all"  VALUE="1">
        <INPUT TYPE="HIDDEN" NAME="cond_basis"   VALUE="Q">
        <INPUT TYPE="HIDDEN" NAME="cond_all"     VALUE="0">
        <INPUT TYPE="HIDDEN" NAME="cond_op"      VALUE="=">
        <INPUT TYPE="HIDDEN" NAME="cond_column"  VALUE="pfid">
        <INPUT TYPE="HIDDEN" NAME="award_op"     VALUE="=">
        <INPUT TYPE="HIDDEN" NAME="award_all"    VALUE="0">
        <INPUT TYPE="HIDDEN" NAME="award_column" VALUE="pfid">
    <INPUT TYPE="HIDDEN" NAME="disjoint_cond_award" VALUE="0">
        <INPUT TYPE="HIDDEN" NAME="disc_type"    VALUE="$">
    <% REM -- buy 2 x for the price of 1---------------------------------------
    Case "3" %>
        <INPUT TYPE="HIDDEN" NAME="promo_type"   VALUE="3">
        <INPUT TYPE="HIDDEN" NAME="promo_description">
        <INPUT TYPE="HIDDEN" NAME="shopper_all"  VALUE="1">
        <INPUT TYPE="HIDDEN" NAME="cond_basis"   VALUE="Q">
        <INPUT TYPE="HIDDEN" NAME="cond_all"     VALUE="0">
        <INPUT TYPE="HIDDEN" NAME="cond_min"     VALUE="2">
        <INPUT TYPE="HIDDEN" NAME="cond_op"      VALUE="=">
        <INPUT TYPE="HIDDEN" NAME="cond_column"  VALUE="pfid">
        <INPUT TYPE="HIDDEN" NAME="award_op"     VALUE="=">
        <INPUT TYPE="HIDDEN" NAME="award_max"    VALUE="1">
        <INPUT TYPE="HIDDEN" NAME="award_all"    VALUE="0">
        <INPUT TYPE="HIDDEN" NAME="award_column" VALUE="pfid">
        <INPUT TYPE="HIDDEN" NAME="award_value">
        <INPUT TYPE="HIDDEN" NAME="disc_type"    VALUE="%">
    <INPUT TYPE="HIDDEN" NAME="disjoint_cond_award" VALUE="0">
        <INPUT TYPE="HIDDEN" NAME="disc_value"   VALUE="100">
    <% REM -- default new mscsPage ------------------------------------------------
    Case else %>
        <INPUT TYPE="HIDDEN" NAME="promo_type"   VALUE="100">
<% End Select '================================================================ %>

<TR>
    <TH ALIGN="LEFT" VALIGN="TOP" WIDTH="130"> Promo name: </TH>
    <TD COLSPAN="3"> 
      <INPUT NAME="promo_name" 
             TYPE="text" 
             SIZE=32
             MAXLENGTH="255"> 
    </TD>
</TR>

<% if PromoType = "100" then %>
    <TR>
        <TH ALIGN="LEFT" VALIGN="TOP"> Description: </TH>
        <TD COLSPAN="3"> 
          <INPUT NAME="promo_description" 
                 TYPE="text" 
                 SIZE=64
                 MAXLENGTH="255"> 
        </TD>
    </TR>
<% end if %>

<TR>
    <TH ALIGN="LEFT" VALIGN="TOP"> Status: </TH>
    <TD WIDTH="130">
        <SELECT NAME="active">
            <% = mscsPage.Option(1, 1) %>  ON
            <% = mscsPage.Option(0, 1) %>  OFF
        </SELECT>
    </TD>
    <TH ALIGN="LEFT" VALIGN="TOP"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Rank: </TH>
    <TD WIDTH="130"> 
        <SELECT NAME="promo_rank">
            <% for count = 10 to 100 step 10 %>
                <% = mscsPage.Option(count, 1) %> <% = count %>
            <% next %>
        </SELECT>
    </TD>
</TR>

<TR>
    <TH ALIGN="LEFT" VALIGN="TOP"> Start date: </TH>
    <TD> 
      <INPUT NAME="date_start" 
             VALUE="<% = FormatDateTime(Now,vbShortDate) %>" 
             TYPE="text" 
             SIZE=12
             MAXLENGTH="12"> 
    </TD>
    <TH ALIGN="LEFT" VALIGN="TOP"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;End date: </TH>
    <TD> 
      <INPUT NAME="date_end"
             VALUE="<% = FormatDateTime(Now + 14,vbShortDate) %>" 
             TYPE="TEXT" 
             SIZE=12
             MAXLENGTH="12"> 
    </TD>
</TR>

<% Select Case PromoType '===============================================
    REM -- buy x get y at z% off-----------------------------------------------
    Case "1" %>
<TR>
    <TH ALIGN="LEFT" VALIGN="TOP"> Buy: </TH>
    <TD COLSPAN="3">
        <INPUT NAME="cond_min" 
               TYPE="text" 
               SIZE=4
               MAXLENGTH="4"
               VALUE=2>
        <SELECT NAME="cond_value">
            <% productOptions %>
        </SELECT>
    </TD>
</TR>

<TR>
    <TH ALIGN="LEFT" VALIGN="TOP"> Get: </TH>
    <TD COLSPAN="3">
        <INPUT NAME="award_max" 
               TYPE="text" 
               SIZE=4
               MAXLENGTH="4"
               VALUE=1>
        <SELECT NAME="award_value">
            <% productOptions %>
        </SELECT>
        at
        <SELECT NAME="disc_value">
            <% = mscsPage.Option(100, 1) %> FREE!
            <% for count = 95 to 5 step -5 %>
                <% = mscsPage.Option(count, 1) %> <% = count %>% OFF&nbsp;
            <% next %>
        </SELECT>
    </TH>
</TR>
    <% REM -- buy x get y at $z off---------------------------------------------
    Case "2" %>
<TR>
    <TH ALIGN="LEFT" VALIGN="TOP"> Buy: </TH>
    <TD COLSPAN="3">
        <INPUT NAME="cond_min" 
               TYPE="text" 
               SIZE=4
               MAXLENGTH="4"
               VALUE=2>
        <SELECT NAME="cond_value">
            <% productOptions %>
        </SELECT>
    </TD>
</TR>

<TR>
    <TH ALIGN="LEFT" VALIGN="TOP"> Get: </TH>
    <TD COLSPAN="3">
        <INPUT NAME="award_max" 
               TYPE="text" 
               SIZE=4
               MAXLENGTH="4"
               VALUE=1>
        <SELECT NAME="award_value">
            <% productOptions %>
        </SELECT>
        at
        <SELECT NAME="disc_value">
            <% for count = 1 to 5 %>
                <% = mscsPage.Option(count, 1) %> $<% = count %> OFF&nbsp;
            <% next %>
            <% for count = 10 to 50 step 5 %>
                <% = mscsPage.Option(count, 1) %> $<% = count %> OFF&nbsp;
            <% next %>
        </SELECT>
    </TH>
</TR>
    <% REM -- buy 2 x for the price of 1---------------------------------------
    Case "3" %>
<TR>
    <TD COLSPAN=4 ALIGN="CENTER" VALIGN="TOP">
        Buy 2
        <SELECT NAME="cond_value">
            <% productOptions %>
        </SELECT>
        get 1 FREE!
    </TD>
</TR>
    <% REM -- default new mscsPage ------------------------------------------------
    Case else %>
<TR>
    <TH ALIGN="LEFT" VALIGN="TOP"> Eligible <BR> Shoppers: </TH>
    <TD COLSPAN="3">
        <TABLE>
            <TR>
            <TD VALIGN="BOTTOM"> 
              <INPUT NAME="shopper_all" 
                     VALUE=0 
                     TYPE="radio"> Shoppers where: 
            </TD>
            <TD VALIGN="BOTTOM">
                <SELECT NAME="shopper_column">
                    <% 
                    for each key in shopperList %>
                        <% = mscsPage.Option(key, 0) %> <% = key %>
                    <% next %>
                </SELECT>
            </TD>
            <TD VALIGN="BOTTOM">
                <SELECT NAME="shopper_op">
                    <OPTION VALUE="<"  > &lt;
                    <OPTION VALUE="<=" > &lt;=
                    <OPTION VALUE="="  SELECTED> =
                    <OPTION VALUE=">=" > &gt;=
                    <OPTION VALUE=">"  > &gt;
                    <OPTION VALUE="<>"  > &lt;&gt;
                </SELECT>
            </TD>
            <TD VALIGN="BOTTOM"> 
              <INPUT NAME="shopper_value" 
                     TYPE="text" 
                     SIZE=16
                     MAXLENGTH="16" > 
            </TD>
            </TR>
            <TR>
            <TD VALIGN="BOTTOM"> 
              <INPUT NAME="shopper_all" 
                     VALUE=1 
                     TYPE="radio" 
                     CHECKED> Any Shopper 
            </TD>
            </TR>
        </TABLE>
    </TD>
</TR>

<TR>
    <TH ALIGN="LEFT" VALIGN="TOP"> Buy: </TH>
    <TD COLSPAN="3">
        <TABLE CELLSPACING=0>
            <TR>
                <TD VALIGN="BOTTOM">
                    <INPUT NAME="cond_min" 
                           TYPE="text" 
                           SIZE=8
                           MAXLENGTH="8"
                           VALUE="1">
                </TD>
                <TD VALIGN="BOTTOM">
                    <SELECT NAME="cond_basis">
                        <OPTION VALUE="Q" SELECTED>  unit(s)
                        <OPTION VALUE="P" >  $ worth
                    </SELECT>
                </TD>
            </TR>
        </TABLE>
        <TABLE>
            <TR>
                <TD VALIGN="BOTTOM"> 
                  <INPUT NAME="cond_all" 
                         VALUE=0 
                         TYPE="radio" 
                         CHECKED> Products where: 
                </TD>
                <TD VALIGN="BOTTOM">
                    <SELECT NAME="cond_column">
                        <% 
                        for each key in productList %>
                            <% = mscsPage.Option(key, 0) %> <% = key %>
                        <%
                        next
                        on error goto 0
                        %>
                    </SELECT>
                </TD>
                <TD VALIGN="BOTTOM">
                    <SELECT NAME="cond_op">
                        <OPTION VALUE="<"  > &lt;
                        <OPTION VALUE="<=" > &lt;=
                        <OPTION VALUE="="  SELECTED> =
                        <OPTION VALUE=">=" > &gt;=
                        <OPTION VALUE=">"  > &gt;
                        <OPTION VALUE="<>"  > &lt;&gt;
                    </SELECT>
                    </TD>
                <TD VALIGN="BOTTOM"> 
                  <INPUT NAME="cond_value"
                         TYPE="text" 
                         SIZE=16
                         MAXLENGTH="16" > 
                </TD>
            </TR>
            <TR>
                <TD VALIGN="BOTTOM"> 
                  <INPUT NAME="cond_all" 
                         VALUE=1 
                         TYPE="radio"> Any Product 
                </TD>
            </TR>
        </TABLE>
    </TD>
</TR>

<TR>
    <TH ALIGN="LEFT" VALIGN="TOP"> Get: </TH>
    <TD COLSPAN="3">
        <TABLE CELLSPACING=0>
        <TR>
        <TD VALIGN="BOTTOM" COLSPAN=4>
          <INPUT NAME="award_max"
                VALUE="1"
                 TYPE="TEXT"
                 SIZE=4
            MAXLENGTH="4"> unit(s)
          <SELECT NAME="disjoint_cond_award">
            <% = mscsPage.Option(0, 1) %> Same as or Different from Buy
            <% = mscsPage.Option(1, 1) %> Different from Buy
          </SELECT>
        </TD>
        </TR>
             
            <TR>
                <TD VALIGN="BOTTOM">
                  <INPUT NAME="award_all" 
                         VALUE=0 
                         TYPE="radio" 
                         CHECKED> Products where: 
                </TD>
                <TD VALIGN="BOTTOM">
                    <SELECT NAME="award_column">
                        <% 
                        for each key in productList %>
                            <% = mscsPage.Option(key, 0) %> <% = key %>
                        <%
                        next
                        %>
                    </SELECT>
                </TD>
                <TD VALIGN="BOTTOM">
                    <SELECT NAME="award_op">
                        <OPTION VALUE="<"  > &lt;
                        <OPTION VALUE="<=" > &lt;=
                        <OPTION VALUE="="  SELECTED> =
                        <OPTION VALUE=">=" > &gt;=
                        <OPTION VALUE=">"  > &gt;
                        <OPTION VALUE="<>"  > &lt;&gt;
                    </SELECT>
                </TD>
                <TD VALIGN="BOTTOM"> 
                  <INPUT NAME="award_value"
                         TYPE="text" 
                         SIZE=16
                         MAXLENGTH="16" > 
                </TD>
            </TR>
            </TR>
                <TD VALIGN="BOTTOM">
                  <INPUT NAME="award_all"
                         VALUE=1 
                         TYPE="radio"> Any Product 
                </TD>
            <TR>
        </TABLE>
    </TD>
</TR>

<TR>
    <TH ALIGN="LEFT" VALIGN="TOP"> At: </TH>
    <TD COLSPAN="3">
        <INPUT NAME="disc_value" 
               TYPE="text" 
               SIZE=8
               MAXLENGTH="8">
        <SELECT NAME="disc_type">
            <OPTION VALUE="$" SELECTED>  $ off
            <OPTION VALUE="%" >  % off
        </SELECT>
    </TD>
</TR>
<% End Select '================================================================ %>

</TABLE>

<BR>
<INPUT TYPE="submit" VALUE="Add Price Promotion">

</FORM>

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
