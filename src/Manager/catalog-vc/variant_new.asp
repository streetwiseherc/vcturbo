<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   VARIANT_NEW.ASP                                                         %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "New Variant" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<TABLE>
    <FORM METHOD="POST"
        ACTION="xt_data_add_update.asp">
        <INPUT TYPE="HIDDEN" NAME="type" VALUE="variant">
        <INPUT TYPE="HIDDEN" NAME="goto" VALUE="variant_list.asp">
        <INPUT TYPE="HIDDEN" NAME="pfid" VALUE="<% = Request("pfid") %>">
    <TR>
        <% REM label:  %> 
        <TD VALIGN=TOP> PF ID: </TD>
        <% REM value:  %> 
        <TD VALIGN=TOP> 
            <STRONG><% = mscsPage.HTMLEncode(Request("pfid")) %></STRONG>
        </TD>
    </TR>
    <TR>
        <% REM label:  %> 
        <TD VALIGN=TOP> Sku: </TD>
        <% REM value:  %> 
        <TD VALIGN=TOP> 
            <INPUT
                TYPE="text"
                SIZE=32
                MAXLENGTH="32"
                NAME="sku"
                VALUE = "">
        </TD>
    </TR>

    <%
    cmdTemp.CommandText = "select attribute_id, attribute_value from vcturbo_product_attribute where pfid = '" & Replace(Request("pfid"),"'", "''") & "' and attribute_index = 0 order by attribute_id"
    Set rsAttr_name = Server.CreateObject("ADODB.Recordset")
    rsAttr_name.Open cmdTemp, , adOpenStatic, adLockReadOnly
    If rsAttr_name.RecordCount > 0 Then %>
        <% While Not rsAttr_name.EOF %>
            <TR>
            <% REM label:  %>
            <TD VALIGN=TOP>
                <% = rsAttr_name("attribute_value").Value %>:
            </TD>
      
            <% REM value:  %>
            <TD VALIGN=TOP>
            <%
            cmdTemp.CommandText = "select * from vcturbo_product_attribute where pfid = '" & Replace(Request("pfid"),"'", "''") & "' and attribute_id = " & rsAttr_name("attribute_id").Value & " and attribute_index > 0 order by attribute_index"
            Set rsAttrs = Server.CreateObject("ADODB.Recordset")
            rsAttrs.Open cmdTemp, , adOpenStatic, adLockReadOnly
            %>
            <SELECT NAME="attribute<% = rsAttr_name("attribute_id").Value %>">
            <% REM = mscsPage.Option(0, 0) %> * None *
            <%

            While Not rsAttrs.EOF %>
                <% = mscsPage.Option(CStr(rsAttrs("attribute_index").Value), 1) %><% = rsAttrs("attribute_value").Value %>
                <% rsAttrs.MoveNext
            Wend %>
            </SELECT>
            </TD>
            </TR>
            <% rsAttr_name.MoveNext
        Wend
    end if
    if rsAttr_name.RecordCount < 5 then
        if rsAttr_name.RecordCount <= 0 then %>
            <INPUT TYPE=HIDDEN NAME="attribute0" VALUE="0">
        <% end if
        if rsAttr_name.RecordCount <= 1 then %>
            <INPUT TYPE=HIDDEN NAME="attribute1" VALUE="0">
        <% end if
        if rsAttr_name.RecordCount <= 2 then %>
            <INPUT TYPE=HIDDEN NAME="attribute2" VALUE="0">
        <% end if
        if rsAttr_name.RecordCount <= 3 then %>
            <INPUT TYPE=HIDDEN NAME="attribute3" VALUE="0">
        <% end if
        if rsAttr_name.RecordCount <= 4 then %>
            <INPUT TYPE=HIDDEN NAME="attribute4" VALUE="0">
        <% end if
    end if
%>
</TABLE>
  
<BR>
<INPUT TYPE="submit" VALUE="Add Variant">
  
</FORM>

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
