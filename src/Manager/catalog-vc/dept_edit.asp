<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   DEPT_EDIT.ASP                                                           %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "Department '" &  Request("name") & "'" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<% REM   body: %>
<% cmdTemp.CommandText = "select dept_id, parent_id, name, description from vcturbo_dept where dept_id = " & Request("dept_id")
Set rsDept = Server.CreateObject("ADODB.Recordset")
rsDept.Open cmdTemp, , adOpenForwardOnly, adLockReadOnly
%>

<FORM METHOD="POST" ACTION="xt_data_add_update.asp">
    
<INPUT TYPE="HIDDEN" NAME="op" VALUE="update">
<INPUT TYPE="HIDDEN" NAME="type" VALUE="department">
<INPUT TYPE="HIDDEN" NAME="goto" VALUE="dept_list.asp">

<TABLE>
    <TR>
        <% REM  label: %> 
        <TD VALIGN=TOP> 
            ID: 
        </TD>
        <% REM  value: %> 
        <TD VALIGN=TOP>
            <INPUT  TYPE="hidden"
                    NAME="dept_id"
                    VALUE="<% = rsDept("dept_id").Value %>">
            <STRONG><% = rsDept("dept_id").Value %></STRONG>
        </TD>
    </TR>

    <TR>
        <% REM  label: %>
        <TD VALIGN=TOP> Parent ID: </TD>

        <% REM  value: %>
        <TD VALIGN=BOTTOM>
        <SELECT NAME="parent_id">
            <% = mscsPage.Option(0, Cstr(rsDept("parent_id").Value)) %> None
            <% 
            cmdTemp.CommandText = "select * from vcturbo_dept where parent_id = 0 order by dept_id"
            Set rsParent = Server.CreateObject("ADODB.Recordset")
            rsParent.Open cmdTemp, , adOpenKeyset, adLockReadOnly
            If rsParent.RecordCount > 0 Then 
                While Not rsParent.EOF %>
                    <% = mscsPage.Option(Cstr(rsParent("dept_id").Value), Cstr(rsDept("parent_id").Value)) %> <% = rsParent("name").Value %>
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
                SIZE = "32"
                NAME = "name"
                MAXLENGTH="255"
                VALUE = "<% = rsDept("name").Value %>">
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
                COLS=40
                ROWS=5
                NAME="description"
                WRAP="virtual"><% = rsDept("description").Value %>
            </TEXTAREA>
        </TD>
    </TR>

</TABLE>

<BR>
* Indicates a required field
<BR>
<BR>

<TABLE>
<TR>
    <TD>
        <INPUT TYPE="submit" VALUE="Update Department">
        </FORM>
    </TD>

    <TD>
        <FORM METHOD="POST" ACTION="confirm.asp">
            <INPUT TYPE="HIDDEN" NAME="confirm_action" VALUE="xt_data_delete.asp">
            <INPUT TYPE="HIDDEN" NAME="confirm_message" VALUE="Delete the department '<% = rsDept("dept_id").Value %>'?">
            <INPUT TYPE="HIDDEN" NAME="type"    VALUE="department">
            <INPUT TYPE="HIDDEN" NAME="goto"    VALUE="dept_list.asp">            
            <INPUT TYPE="HIDDEN" NAME="dept_id" VALUE="<% = rsDept("dept_id").Value %>">
            <INPUT TYPE="SUBMIT" VALUE="Delete Department...">
        </FORM>
    </TD>
</TR>
</TABLE>

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->

