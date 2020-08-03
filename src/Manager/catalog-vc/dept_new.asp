<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   DEPT_NEW.ASP                                                            %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "New Department" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<FORM METHOD="POST"
    ACTION="xt_data_add_update.asp">
    <INPUT TYPE="HIDDEN" NAME="type" VALUE="department">
    <INPUT TYPE="HIDDEN" NAME="goto" VALUE="dept_list.asp">

<TABLE>
    <TR>
        <% REM  label: %> 
        <TD VALIGN=TOP> 
            * ID: 
        </TD>
        <% REM  value: %> 
        <TD VALIGN=TOP>
            <INPUT  TYPE="text"
                    SIZE = 32
                    NAME="dept_id">
        </TD>
    </TR>

    <TR>
        <% REM  label: %>
        <TD VALIGN=TOP> Parent ID: </TD>

        <% REM  value: %>
        <TD VALIGN=BOTTOM>
        <SELECT NAME="parent_id">
            <% = mscsPage.Option(0, 0) %> None
            <% 
            cmdTemp.CommandText = "select * from vcturbo_dept where parent_id = 0 order by dept_id"
            Set rsParent = Server.CreateObject("ADODB.Recordset")
            rsParent.Open cmdTemp, , adOpenKeyset, adLockReadOnly
            If rsParent.RecordCount > 0 Then 
                While Not rsParent.EOF %>
                    <% = mscsPage.Option(Cstr(rsParent("dept_id").Value), 0) %> <% = rsParent("name").Value %>
                    <% rsParent.MoveNext
                Wend
            End If
            %>
        </SELECT>
        </TD>
    </TR>

    <TR>
        <% REM  label: %>
        <TD VALIGN=TOP>
            * Name:
        </TD>

        <% REM  value: %>
        <TD VALIGN=TOP>
            <INPUT
                TYPE = "text" 
                SIZE = 32
                MAXLENGTH ="255"
                NAME = "name">
        </TD>
    </TR>


    <TR>
        <% REM label:  %>
        <TD VALIGN=TOP>
            Description:
        </TD>

        <% REM value:  %>
        <TD VALIGN=TOP>
            <TEXTAREA
                COLS="40"
                ROWS="5"
                NAME="description"
                WRAP = "virtual">
            </TEXTAREA>
        </TD>
    </TR>

</TABLE>

<BR>
* Indicates a required field
<BR>
<BR>
<INPUT TYPE="submit" VALUE="Add Department">

</FORM>

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->

