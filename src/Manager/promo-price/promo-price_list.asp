<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   PROMO-PRICE_LIST.ASP                                                    %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "Price Promotions" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->


<A HREF="promo-price_new.asp"><H2> Add New Price Promotion</H2> </A>
<FORM METHOD="GET" ACTION="promo-price_new.asp">
    <H2>Add Special Promotion:      
        <SELECT NAME="promo_type">
            <% = mscsPage.Option(1, 1) %> buy x get y at z% off
            <% = mscsPage.Option(2, 1) %> buy x get y at $z off
            <% = mscsPage.Option(3, 1) %> buy 2 x for the price of 1&nbsp;
        </SELECT>
    <INPUT TYPE="SUBMIT" VALUE="Add"></H2>
</FORM>

<%
function ListRow() %>
    <TD VALIGN=TOP ALIGN=CENTER> <% = RowCount %>            </TD>
    <TD VALIGN=TOP ALIGN=LEFT  > <A HREF="<% = "promo-price_edit.asp?" & mscsPage.URLArgs("promo_name", rsList("promo_name").Value) %>"> <% = rsList("promo_name").Value %> </A> </TD>
    <TD VALIGN=TOP ALIGN=LEFT  > <% = rsList("promo_description") %> </TD>
    <TD VALIGN=TOP ALIGN=LEFT  > <% = rsList("promo_rank") %>        </TD>
    <TD VALIGN=TOP ALIGN=LEFT  > 
        <% If rsList("active") Then %>Yes<% Else %>No<% End If %>
    </TD>
    <TD VALIGN=TOP ALIGN=LEFT  > <% = FormatDateTime(rsList("date_start"),vbShortDate) %>         </TD>
    <TD VALIGN=TOP ALIGN=LEFT  > <% = FormatDateTime(rsList("date_end"),vbShortDate) %>           </TD>
<% end function 
function ListHeader() %>
    <TH ALIGN=LEFT VALIGN=BOTTOM> #     </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Name  </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Desc  </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Rank  </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> On?   </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> Start </TH>
    <TH ALIGN=LEFT VALIGN=BOTTOM> End   </TH>
<% end function
ListSQL = "SELECT * FROM vcturbo_promo_price ORDER BY active DESC, promo_rank, date_start, promo_name"
ListNoRows = "No promotions found"
%>
<!--#INCLUDE FILE="include/list.asp" -->

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
