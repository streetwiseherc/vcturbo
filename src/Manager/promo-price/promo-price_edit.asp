<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   PROMO-PRICE_EDIT.ASP                                                    %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% promo_name = mscsPage.RequestString("promo_name") %>

<% REM This code fills in the data to be used for the shopper and product selection boxes


    REM Fields available for product selection
    set productList = Server.CreateObject("Commerce.SimpleList")
    productList.Add("sku")
    productList.Add("name")
    productList.Add("list_price")
    productList.Add("pfid")

    REM Fields available for shopper selection
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
    cmdTemp.CommandText = "SELECT * FROM vcturbo_promo_price WHERE promo_name =  '" &  Replace(promo_name, "'", "''") & "'"
    Set rsPromo = Server.CreateObject("ADODB.Recordset")
    rsPromo.Open cmdTemp, , adOpenKeyset, adLockReadOnly

    cmdTemp.CommandText = "SELECT pfid, name FROM vcturbo_product_family ORDER BY name"
    Set rsProduct = Server.CreateObject("ADODB.Recordset")
    rsProduct.Open cmdTemp, , adOpenKeyset, adLockReadOnly

    sub productOptions(default)
        rsProduct.Movefirst
        While Not rsProduct.EOF
            Response.write mscsPage.Option(Cstr(rsProduct("pfid").Value), default) & " " & rsProduct("name").Value & "&nbsp;" & vbCr
            rsProduct.MoveNext
        Wend
    end sub


    PromoType = mscsPage.RequestString("promo_type")
    if IsNull(PromoType) then
        PromoType = rsPromo("promo_type")
    end if
%>

<% REM   header: %>
<% pageTitle = "Edit Price Promo '" &  mscsPage.HTMLEncode(promo_name) & "'" %>

<HTML>
<HEAD>
    <TITLE> <% = pageTitle %> </TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>

<BODY TOPMARGIN="<% = top_margin %>" LEFTMARGIN="<% = left_margin %>" BGCOLOR="<% = bgcolor %>" TEXT="<% = text %>" LINK="<% = color_link %>" ALINK="<% = color_alink %>" VLINK="<% = color_vlink %>">

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<TABLE BORDER="1" CELLPADDING="5" BORDERCOLOR="#000000">

<% Select Case PromoType '===============================================
    REM -- buy x get y at z% off-----------------------------------------------
    Case "1" %>
        <SCRIPT LANGUAGE="JavaScript">
        <!--
            function promoform_onsubmit() {
        choices = new Array(5)
                choices[0] = document.promoform.cond_value.options[document.promoform.cond_value.selectedIndex].value;
                choices[1] = document.promoform.award_value.options[document.promoform.award_value.selectedIndex].value;
        choices[2] = document.promoform.disc_value.options[document.promoform.disc_value.selectedIndex].value;
        choices[3] = document.promoform.cond_min.value;
        choices[4] = document.promoform.award_max.value;
        description = "Buy %3 '%0' get %4 '%1' at %2 % OFF!";
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
        choices = new Array(5)
                choices[0] = document.promoform.cond_value.options[document.promoform.cond_value.selectedIndex].value;
                choices[1] = document.promoform.award_value.options[document.promoform.award_value.selectedIndex].value;
                choices[2] = document.promoform.disc_value.options[document.promoform.disc_value.selectedIndex].value;
        choices[3] = document.promoform.cond_min.value;
        choices[4] = document.promoform.award_max.value;
        description = "Buy %3 '%0' get %4 '%1' at $2 OFF!";
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
        <INPUT TYPE="HIDDEN" NAME="op" VALUE="update">

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
        <INPUT TYPE="HIDDEN" NAME="disjoint_cond_award" VALUE="0">
        <INPUT TYPE="HIDDEN" NAME="disc_type"    VALUE="%">
        <INPUT TYPE="HIDDEN" NAME="disc_value"   VALUE="100">
    <% REM -- default new page ------------------------------------------------
    Case else %>
        <INPUT TYPE="HIDDEN" NAME="promo_type"   VALUE="100">
<% End Select '================================================================ %>

<TR>
    <TH ALIGN="LEFT" VALIGN="TOP" WIDTH="130"> Promo name: </TH>
    <TD COLSPAN="3"> 
        <INPUT  TYPE="hidden"
                NAME="promo_name"
                VALUE="<% = rsPromo("promo_name").Value %>">
        <STRONG><% = rsPromo("promo_name").Value %></STRONG>
    </TD>
</TR>

<% if PromoType = "100" then %>
    <TR>
        <TH ALIGN="LEFT" VALIGN="TOP"> Description: </TH>
        <TD COLSPAN="3"> 
          <INPUT NAME="promo_description" 
                 TYPE="text" 
                 SIZE=64
                 VALUE="<% = rsPromo("promo_description").Value %>"> 
        </TD>
    </TR>
<% end if %>

<TR>
    <TH ALIGN="LEFT" VALIGN="TOP"> Status: </TH>
    <TD WIDTH="130">
        <SELECT NAME="active">
            <% = mscsPage.Option(1, rsPromo("active").Value) %>  ON&nbsp;
            <% = mscsPage.Option(0, rsPromo("active").Value) %>  OFF&nbsp;
        </SELECT>
    </TD>
    <TH ALIGN="LEFT" VALIGN="TOP"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Rank: </TH>
    <TD WIDTH="130"> 
        <SELECT NAME="promo_rank">
            <% for count = 10 to 100 step 10 %>
                <% = mscsPage.Option(Cstr(count), rsPromo("promo_rank").Value) %> <% = count %>
            <% next %>
        </SELECT>
    </TD>
</TR>

<TR>
    <TH ALIGN="LEFT" VALIGN="TOP"> Start date: </TH>
    <TD> 
      <INPUT NAME="date_start" 
             VALUE="<% = FormatDateTime(rsPromo("date_start").Value,vbShortDate) %>" 
             TYPE="text" 
             SIZE=12
             MAXLENGTH="12"> 
    </TD>
    <TH ALIGN="LEFT" VALIGN="TOP"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;End date: </TH>
    <TD> 
      <INPUT NAME="date_end"
             VALUE="<% = FormatDateTime(rsPromo("date_end").Value,vbShortDate) %>" 
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
               VALUE="<% = rsPromo("cond_min").Value %>">
        <SELECT NAME="cond_value">
            <% productOptions rsPromo("cond_value").Value %>
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
               VALUE="<% = rsPromo("award_max").Value %>">
        <SELECT NAME="award_value">
            <% productOptions rsPromo("award_value").Value %>
        </SELECT>
        at
        <SELECT NAME="disc_value">
            <% = mscsPage.Option(100, 1) %> FREE!
            <% for count = 95 to 5 step -5 %>
                <% = mscsPage.Option(CLng(count), CLng(rsPromo("disc_value").Value)) %> <% = count %> % OFF&nbsp;
            <% next %>
        </SELECT>
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
               VALUE="<% = rsPromo("cond_min").Value %>">
        <SELECT NAME="cond_value">
            <% productOptions rsPromo("cond_value").Value %>
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
               VALUE="<% = rsPromo("award_max").Value %>">
        <SELECT NAME="award_value">
            <% productOptions rsPromo("award_value").Value %>
        </SELECT>
        at
        <SELECT NAME="disc_value">
            <% for count = 1 to 5 %>
                <% = mscsPage.Option(count, rsPromo("disc_value").Value) %> $<% = count %> OFF&nbsp;
            <% next %>
            <% for count = 10 to 50 step 5 %>
                <% = mscsPage.Option(count, rsPromo("disc_value").Value) %> $<% = count %> OFF&nbsp;
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
            <% productOptions rsPromo("cond_value").Value %>
        </SELECT>
        get 1 FREE!
    </TD>
</TR>
    <% REM -- default new page ------------------------------------------------
    Case else %>
<TR>
    <TH ALIGN="LEFT" VALIGN="TOP"> Eligible <BR> Shoppers: </TH>
    <TD COLSPAN="3">
        <TABLE>
            <TR>
            <TD VALIGN="BOTTOM"> 
              <INPUT NAME="shopper_all" 
                     VALUE=0 
                     TYPE="radio"
                 <% = mscsPage.Check(Cint(rsPromo("shopper_all").Value) = 0) %>> Shoppers where: 
            </TD>
            <TD VALIGN="BOTTOM">
                <SELECT NAME="shopper_column">
                    <% 
                    if IsNull(rsPromo("shopper_column").Value) then
                        shopper_column = ""
                    else
                        shopper_column = rsPromo("shopper_column").Value
                    end if
                    for each key in shopperList %>
                        <% = mscsPage.Option(LCase(key), shopper_column) %> <% = key %>
                    <%
                    next
                    %>
                </SELECT>
            </TD>
            <TD VALIGN="BOTTOM">
                <% 
                if IsNull(rsPromo("shopper_op").Value) then
                    shopper_op = "="
                else
                    shopper_op = rsPromo("shopper_op").Value
                end if
                %>
                <SELECT NAME="shopper_op">
                    <% = mscsPage.Option("<",  Cstr(shopper_op)) %> &lt;
                    <% = mscsPage.Option("<=", Cstr(shopper_op)) %> &lt;=
                    <% = mscsPage.Option("=",  Cstr(shopper_op)) %> =
                    <% = mscsPage.Option(">=", Cstr(shopper_op)) %> &gt;=
                    <% = mscsPage.Option(">",  Cstr(shopper_op)) %> &gt;
                    <% = mscsPage.Option("<>", Cstr(shopper_op)) %> &lt;&gt;
                </SELECT>
            </TD>
            <TD VALIGN="BOTTOM"> 
              <INPUT NAME="shopper_value" 
                     TYPE="text" 
                     SIZE=16
                     MAXLENGTH="16"
                     VALUE="<% = rsPromo("shopper_value").Value %>"> 
            </TD>
            </TR>
            <TR>
            <TD VALIGN="BOTTOM"> 
              <INPUT NAME="shopper_all" 
                     VALUE=1 
                     TYPE="radio"
                 <% = mscsPage.Check(Cint(rsPromo("shopper_all").Value) = 1) %>> Any Shopper 
            </TD>
            </TR>
        </TABLE>
    </TD>
</TR>

<TR>
    <TH ALIGN="LEFT" VALIGN="TOP"> Buy: </TH>
    <TD COLSPAN="3">
        <TABLE CELLSPACING="0">
            <TR>
                <TD VALIGN="BOTTOM">
                    <INPUT NAME="cond_min" 
                           TYPE="text" 
                           SIZE=8
                           MAXLENGTH="8"
                           VALUE="<% if rsPromo("cond_basis").Value = "P" then %><% = MSCSDataFunctions.Money(Clng(rsPromo("cond_min").Value)) %><% else %><% = rsPromo("cond_min").Value %><% end if %>" >
                </TD>
                <TD VALIGN="BOTTOM">
                    <% 
                    if IsNull(rsPromo("cond_basis").Value) then
                        cond_basis = ""
                    else
                        cond_basis = rsPromo("cond_basis").Value
                    end if
                    %>
                    <SELECT NAME="cond_basis">
                        <% = mscsPage.Option("Q", Cstr(cond_basis)) %> unit(s)
                        <% = mscsPage.Option("P", Cstr(cond_basis)) %> $ worth
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
                         <% = mscsPage.Check(Cbool(rsPromo("cond_all").Value) = 0) %>> Products where: 
                </TD>
                <TD VALIGN="BOTTOM">
                    <SELECT NAME="cond_column">
                        <% 
                        cond_column = rsPromo("cond_column").Value
                        for each key in productList %>
                            <% = mscsPage.Option(Cstr(key), cond_column) %> <% = key %>
                        <%
                        next
                        %>
                    </SELECT>
                </TD>
                <TD VALIGN="BOTTOM">
                    <% 
                    if IsNull(rsPromo("cond_op").Value) then
                        cond_op = "="
                    else
                        cond_op = rsPromo("cond_op").Value
                    end if
                    %>
                    <SELECT NAME="cond_op">
                        <% = mscsPage.Option("<",  Cstr(cond_op)) %> &lt;
                        <% = mscsPage.Option("<=", Cstr(cond_op)) %> &lt;=
                        <% = mscsPage.Option("=",  Cstr(cond_op)) %> =
                        <% = mscsPage.Option(">=", Cstr(cond_op)) %> &gt;=
                        <% = mscsPage.Option(">",  Cstr(cond_op)) %> &gt;
                        <% = mscsPage.Option("<>", Cstr(cond_op)) %> &lt;&gt;
                    </SELECT>
                    </TD>
                <TD VALIGN="BOTTOM"> 
                  <INPUT NAME="cond_value"
                         TYPE="text" 
                         SIZE=16
                         MAXLENGTH="16"
                         VALUE="<% = rsPromo("cond_value").Value %>"> 
                </TD>
            </TR>
            <TR>
                <TD VALIGN="BOTTOM"> 
                  <INPUT NAME="cond_all" 
                         VALUE=1 
                         TYPE="radio" 
                         <% = mscsPage.Check(Cbool(rsPromo("cond_all").Value = 1)) %>> Any Product 
                </TD>
            </TR>
        </TABLE>
    </TD>
</TR>

<TR>
    <TH ALIGN="LEFT" VALIGN="TOP"> Get: </TH>
    <TD COLSPAN="3">
        <TABLE CELLSPACING="0">
        <TR>
        <TD VALIGH="BOTTOM" COLSPAN=4>
          <INPUT NAME="award_max"
             VALUE="<% = rsPromo("award_max").Value %>"
             TYPE="TEXT"
             SIZE=4
             MAXLENGTH="4"> unit(s)
          <SELECT NAME="disjoint_cond_award">
            <% = mscsPage.Option(0, rsPromo("disjoint_cond_award").Value) %> Same as or Different from Buy
            <% = mscsPage.Option(1, rsPromo("disjoint_cond_award").Value) %> Different from Buy
          </SELECT>
        </TD>
        </TR>

            <TR>
                <TD VALIGN="BOTTOM">
                  <INPUT NAME="award_all" 
                         VALUE=0 
                         TYPE="radio" 
                         <% = mscsPage.Check(Cbool(rsPromo("award_all").Value) = 0) %>> Products where: 
                </TD>
                <TD VALIGN="BOTTOM">
                    <SELECT NAME="award_column">
                        <% 
                        award_column = rsPromo("award_column").Value
                        on error resume next
                        for each key in productList %>
                            <% = mscsPage.Option(key, award_column) %> <% = key %>
                        <%
                        next
                        %>
                    </SELECT>
                </TD>
                <TD VALIGN="BOTTOM">
                    <% 
                    if IsNull(rsPromo("award_op").Value) then
                        award_op = "="
                    else
                        award_op = rsPromo("award_op").Value
                    end if
                    %>
                    <SELECT NAME="award_op">
                        <% = mscsPage.Option("<",  Cstr(award_op)) %> &lt;
                        <% = mscsPage.Option("<=", Cstr(award_op)) %> &lt;=
                        <% = mscsPage.Option("=",  Cstr(award_op)) %> =
                        <% = mscsPage.Option(">=", Cstr(award_op)) %> &gt;=
                        <% = mscsPage.Option(">",  Cstr(award_op)) %> &gt;
                        <% = mscsPage.Option("<>", Cstr(award_op)) %> &lt;&gt;
                    </SELECT>
                </TD>
                <TD VALIGN="BOTTOM"> 
                  <INPUT NAME="award_value"
                         TYPE="text" 
                         SIZE=16
                         MAXLENGTH="16"
                         VALUE="<% = rsPromo("award_value").Value %>"> 
                </TD>
            </TR>
            </TR>
                <TD VALIGN="BOTTOM">
                  <INPUT NAME="award_all"
                         VALUE=1 
                         TYPE="radio" 
                         <% = mscsPage.Check(rsPromo("award_all").Value = 1) %>> Any Product 
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
               MAXLENGTH="8"
               VALUE="<% if rsPromo("disc_type").Value = "$" then %><% = MSCSDataFunctions.Money(Clng(rsPromo("disc_value").Value)) %><% else %><% = rsPromo("disc_value").Value %><% end if %>">
        <SELECT NAME="disc_type">
            <% = mscsPage.Option("$", rsPromo("disc_type")) %> $ off
            <% = mscsPage.Option("%", rsPromo("disc_type")) %> % off
        </SELECT>
    </TD>
</TR>
<% End Select '================================================================ %>

</TABLE>

<BR>
<% if PromoType = "1" or PromoType = "2" or PromoType = "3" then %>
    <FONT SIZE="-1">Click to edit <A HREF="promo-price_edit.asp?promo_type=100&promo_name=<% = mscsPage.URLEncode(promo_name) %>">Advanced Attributes<A> of the promotion</FONT>
<% end if %>
<BR>
<BR>

<TABLE>
<TR>
<TD>
    <INPUT TYPE="submit" VALUE="Update Price Promotion">
    </FORM>
</TD>

<TD>
    <FORM METHOD="POST" ACTION="confirm.asp">
        <INPUT TYPE="HIDDEN" NAME="confirm_message" VALUE="Delete the price promotion '<%= rsPromo("promo_name").Value %>'?">
        <INPUT TYPE="HIDDEN" NAME="confirm_action" VALUE="xt_data_delete.asp">
        <INPUT TYPE="HIDDEN" NAME="promo_name" VALUE="<% = rsPromo("promo_name").Value %>">
        <INPUT TYPE="SUBMIT" VALUE="Delete Price Promotion...">
    </FORM>
</TD>
</TR>
</TABLE>

<% REM   footer: %>
<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
