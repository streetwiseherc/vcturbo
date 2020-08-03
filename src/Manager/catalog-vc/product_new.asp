<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   PRODUCT_NEW.ASP                                                         %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "New Product" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<FORM METHOD="POST"
    ACTION="xt_data_add_update.asp">
    <INPUT TYPE="HIDDEN" NAME="type" VALUE="product">
    <INPUT TYPE="HIDDEN" NAME="goto" VALUE="product_list.asp">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
    <TR>
        <% REM label:  %>
        <TD VALIGN=TOP>
            * PFId:
        </TD>

        <% REM value:  %>
        <TD VALIGN=TOP colspan=2>
            <INPUT  TYPE="text"
                    SIZE=32
                    MAXLENGTH=32
                    NAME="pfid">
        </TD>
    </TR>
    <TR>
        <% REM label:  %>
        <TD VALIGN=TOP>
            * Name:
        </TD>

        <% REM value:  %>
        <TD VALIGN=TOP colspan=2>
            <INPUT
                TYPE="text"
                SIZE=32
                MAXLENGTH="255"
                NAME="name">
        </TD>
    </TR>
    <TR>
        <% REM label:  %>
        <TD VALIGN=TOP>
            Short Description:
        </TD>

        <% REM value:  %>
        <TD VALIGN=TOP colspan=2>
            <TEXTAREA
                COLS=36
                ROWS=5
                NAME="short_description"
                WRAP = "virtual"></TEXTAREA>
        </TD>
    </TR>
    <TR>
        <% REM label:  %>
        <TD VALIGN=TOP>
            Long Description:
        </TD>

        <% REM value:  %>
        <TD VALIGN=TOP colspan=2>
            <TEXTAREA
                COLS=36
                ROWS=5
                NAME="long_description"
                WRAP = "virtual"></TEXTAREA>
        </TD>
    </TR>              
    <TR>
        <% REM label:  %>
        <TD VALIGN=TOP>
            * Price:
        </TD>

        <% REM value:  %>
        <TD VALIGN=TOP colspan=2>
            <INPUT
                TYPE="text"
                SIZE=32
                MAXLENGTH=32
                NAME="list_price">
        </TD>
    </TR>
    <TR>
        <% REM label:  %>
        <TD VALIGN=TOP>
            Dept:
        </TD>

        <% REM value:  %>
        <TD VALIGN=TOP colspan=2>
            <SELECT NAME="dept_id">
                <% 
                cmdTemp.CommandText = "select dept_id, name from vcturbo_dept where parent_id <> 0"
                Set rsDept = Server.CreateObject("ADODB.Recordset")
                rsDept.Open cmdTemp, , adOpenKeyset, adLockReadOnly
                If rsDept.RecordCount > 0 Then 
                    While Not rsDept.EOF %>
                        <% = mscsPage.Option(Cstr(rsDept("dept_id").Value), 0) %> <% = rsDept("name").Value %>
                        <% rsDept.MoveNext
                    Wend
                End If
                %>
            </SELECT>
        </TD>
    </TR>
    <TR>
        <% REM label:  %>
        <TD VALIGN=TOP>
            Monogramable?:
        </TD>

        <% REM value:  %>
        <TD VALIGN=TOP colspan=2>
            <INPUT
                TYPE="checkbox"
                VALUE=1
                NAME="monogramable">Yes
        </TD>

    </TR>
    <TR>
        <% REM label:  %>
        <TD VALIGN=TOP>
            Image File:
        </TD>

        <% REM value:  %>
        <TD VALIGN=TOP colspan=2>
            <INPUT
                TYPE="text"
                SIZE=28
                MAXLENGTH=28
                NAME="image_filename">
        </TD>

    </TR>
</TABLE>

<BR>
* Indicates a required field
<BR>
<BR>
<INPUT TYPE="submit" VALUE="Add Product">
</FORM>

<BR>
<H5>
NOTE: Remember that you must add at least one variant
for each product in order for the store to work properly!
</H5>

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
