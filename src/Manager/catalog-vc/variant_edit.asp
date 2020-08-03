<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   VARIANT_EDIT.ASP                                                        %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "Edit Variant '" &  Request("sku") & "'" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<%
cmdTemp.CommandText = "select * from vcturbo_product_variant where sku = " & Request("sku")
Set rsPvariant = Server.CreateObject("ADODB.Recordset")
rsPvariant.Open cmdTemp, , adOpenStatic, adLockReadOnly

if rsPvariant.RecordCount <> 1 then
    Response.Write("<H2>Invalid Variant</H2>")
else
    
%>

<A HREF="product_edit.asp?pfid=<% = mscsPage.URLEncode(rsPvariant("pfid").Value) %>">
<H2>
PFid: <% = mscsPage.HTMLEncode(rsPvariant("pfid").Value) %>
</H2>
</A>

<FORM METHOD=POST ACTION="xt_data_add_update.asp">
    <INPUT TYPE="HIDDEN" NAME="op" VALUE="update">
    <INPUT TYPE="HIDDEN" NAME="type" VALUE="variant">
    <INPUT TYPE="HIDDEN" NAME="goto" VALUE="variant_list.asp">
<TABLE>
    <TR>
        <% REM label:  %> 
        <TD VALIGN=TOP>SKU:</TD>
  
        <% REM value:  %>
        <TD VALIGN=TOP>
            <INPUT
                TYPE="hidden"
                NAME="sku"
                VALUE = "<% = rsPvariant("sku").Value %>">
            <STRONG><% = rsPvariant("sku").Value %></STRONG>
        </TD>
    </TR>

    <TR>
        <% REM label:  %> 
        <TD VALIGN=TOP>PF ID:</TD>
  
        <% REM value:  %>
        <TD VALIGN=TOP>
            <INPUT
                TYPE="hidden"
                NAME="pfid"
                VALUE = "<% = rsPvariant("pfid").Value %>">
            <STRONG><% = rsPvariant("pfid").Value %></STRONG>
        </TD>
    </TR>

    <%
    cmdTemp.CommandText = "select attribute_id, attribute_value from vcturbo_product_attribute where pfid = '" & Replace(rsPvariant("pfid"),"'", "''") & "' and attribute_index = 0 order by attribute_id"
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
            <SELECT NAME="attribute<% = (Cint(rsAttr_name("attribute_id").Value)) %>">
            <% if Cint(rsAttr_name("attribute_id").Value) = 0 then %>
                <% REM = mscsPage.Option(0, CStr(rsPvariant("attribute0").Value)) %> * None *
            <% end if
            if Cint(rsAttr_name("attribute_id").Value) = 1 then %>
                <% REM = mscsPage.Option(0, CStr(rsPvariant("attribute1").Value)) %> * None *
            <% end if
            if Cint(rsAttr_name("attribute_id").Value) = 2 then %>
                <% REM = mscsPage.Option(0, CStr(rsPvariant("attribute2").Value)) %> * None *
            <% end if
            if Cint(rsAttr_name("attribute_id").Value) = 3 then %>
                <% REM = mscsPage.Option(0, CStr(rsPvariant("attribute3").Value)) %> * None *
            <% end if
            if Cint(rsAttr_name("attribute_id").Value) = 4 then %>
                <% REM = mscsPage.Option(0, CStr(rsPvariant("attribute4").Value)) %> * None *
            <% end if %>
            <%
            cmdTemp.CommandText = "select * from vcturbo_product_attribute where pfid = '" & Replace(rsPvariant("pfid"),"'", "''") & "' and attribute_id = " & rsAttr_name("attribute_id") & " and attribute_index > 0 order by attribute_index"
            Set rsAttrs = Server.CreateObject("ADODB.Recordset")
            rsAttrs.Open cmdTemp, , adOpenStatic, adLockReadOnly
            While Not rsAttrs.EOF
                if Cint(rsAttr_name("attribute_id").Value) = 0 then %>
                    <% = mscsPage.Option(CStr(rsAttrs("attribute_index").Value), CStr(rsPvariant("attribute0").Value)) %><% = rsAttrs("attribute_value").Value %>
                <% end if 
                if Cint(rsAttr_name("attribute_id").Value) = 1 then %>
                    <% = mscsPage.Option(CStr(rsAttrs("attribute_index").Value), CStr(rsPvariant("attribute1").Value)) %><% = rsAttrs("attribute_value").Value %>
                <% end if 
                 if Cint(rsAttr_name("attribute_id").Value) = 2 then %>
                    <% = mscsPage.Option(CStr(rsAttrs("attribute_index").Value), CStr(rsPvariant("attribute2").Value)) %><% = rsAttrs("attribute_value").Value %>
                <% end if
                if Cint(rsAttr_name("attribute_id").Value) = 3 then %>
                    <% = mscsPage.Option(CStr(rsAttrs("attribute_index").Value), CStr(rsPvariant("attribute3").Value)) %><% = rsAttrs("attribute_value").Value %>
                <% end if
                if Cint(rsAttr_name("attribute_id").Value) = 4 then %>
                    <% = mscsPage.Option(CStr(rsAttrs("attribute_index").Value), CStr(rsPvariant("attribute4").Value)) %><% = rsAttrs("attribute_value").Value %>
                <% end if
                rsAttrs.MoveNext
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
    end if %>
</TABLE>
  
<BR>
<TABLE>
    <TR>
        <TD VALIGN=TOP>
            <% if rsAttr_name.RecordCount > 0 then %>
                <INPUT TYPE="submit" 
                       VALUE="Update Variant">
            <% end if %>
        </TD>
</FORM>
        <TD VALIGN=TOP>
            <FORM METHOD=POST ACTION="confirm.asp">
                <INPUT TYPE="HIDDEN" NAME="confirm_action" VALUE="xt_data_delete.asp">
                <INPUT TYPE="HIDDEN" NAME="confirm_message" VALUE="Delete the variant '<%= rsPvariant("sku").Value %>'?">
                <INPUT TYPE="HIDDEN" NAME="type" VALUE="variant">
                <INPUT TYPE="HIDDEN" NAME="goto" VALUE="variant_list.asp">
                <INPUT TYPE="HIDDEN" NAME="pfid" VALUE="<% = rsPvariant("pfid").Value %>">
                <INPUT TYPE="HIDDEN" NAME="sku"  VALUE="<% = rsPvariant("sku").Value %>">
                <INPUT TYPE="submit" VALUE="Delete Variant...">
            </FORM>
        </TD>
    </TR>
</TABLE>

<% end if %>
<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
