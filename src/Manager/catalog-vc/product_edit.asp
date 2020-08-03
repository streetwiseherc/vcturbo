<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="include/manager.asp" -->
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   PRODUCT_EDIT.ASP                                                        %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% pageTitle = "Edit Product '" &  Request("pfid") & "'" %>

<!--#INCLUDE FILE="include/mgmt_header.asp" -->

<%
cmdTemp.CommandText = "SELECT * FROM vcturbo_product_family where pfid = '" & Replace(Request("pfid"), "'", "''") & "'"
Set rsProduct = Server.CreateObject("ADODB.Recordset")
rsProduct.Open cmdTemp, , adOpenKeyset, adLockReadOnly
%>

<BR>

<% 
if rsProduct.EOF then
    Response.Write("<I>Product no longer exists in store</I>")
    Response.End

end if
%>

<BR>

<TABLE BORDER=0 CELLPADDING=5>

<FORM METHOD="POST" ACTION="xt_data_add_update.asp">
<INPUT TYPE="HIDDEN" NAME="op" VALUE="update">
<INPUT TYPE="HIDDEN" NAME="type" VALUE="product">
<INPUT TYPE="HIDDEN" NAME="goto" VALUE="product_list.asp">

    <TR>
        <TD VALIGN=TOP>

            <TABLE WIDTH=300 CELLPADDING=5>
                <TR>
                    <% REM label:  %>
                    <TD VALIGN=TOP>
                        PFId:
                    </TD>
              
                    <% REM value:  %>
                    <TD VALIGN=TOP colspan=2>
                        <INPUT  TYPE="hidden"
                                NAME="pfid"
                                VALUE="<% = rsProduct("pfid").Value %>">
                        <STRONG><% = rsProduct("pfid").Value %></STRONG>
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
                            NAME="name"
                            VALUE = "<% = rsProduct("name").Value %>">
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
                            WRAP = "virtual"><% = rsProduct("short_description").Value %></TEXTAREA>
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
                            WRAP = "virtual"><% = rsProduct("long_description").Value %></TEXTAREA>
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
                            MAXLENGTH="32"
                            NAME="list_price"
                            VALUE = "<% = MSCSDataFunctions.Money(CLng(rsProduct("list_price").Value)) %>">
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
                            cmdTemp.CommandText = "select * from vcturbo_dept"
                            Set rsDept = Server.CreateObject("ADODB.Recordset")
                            rsDept.Open cmdTemp, , adOpenKeyset, adLockReadOnly
                            If rsDept.RecordCount > 0 Then 
                                While Not rsDept.EOF %>
                                    <% = mscsPage.Option(Cstr(rsDept("dept_id").Value), Cstr(rsProduct("dept_id").Value)) %> <% = rsDept("name").Value %>
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
                            NAME="monogramable"
                            <% if Not IsNull(rsProduct("monogramable").Value) then %><% = mscsPage.Check(CBool(rsProduct("monogramable").Value = "1")) %><% end if %>>Yes
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
                            MAXLENGTH="28"
                            NAME="image_filename"
                            VALUE = "<% = rsProduct("image_filename").Value %>">
                    </TD>
              
                </TR>
</TABLE>
<BR>
* Indicates a required field
<BR>
<BR>
              
<TABLE> 
                <TR>
                    <TD></TD>
                    <TD valign=top>
                        <INPUT TYPE="submit" 
                               VALUE="Update Product">
                    </TD>
                
</FORM>
                    <TD valign=top>
                        <FORM METHOD="POST" ACTION="confirm.asp">
                            <INPUT TYPE="HIDDEN" NAME="confirm_action" VALUE="xt_data_delete.asp">
                            <INPUT TYPE="HIDDEN" NAME="confirm_message" VALUE="Delete the product '<%= rsProduct("pfid").Value %>'?">
                            <INPUT TYPE="HIDDEN" NAME="type" VALUE="product">
                            <INPUT TYPE="HIDDEN" NAME="goto" VALUE="product_list.asp">
                            <INPUT TYPE="HIDDEN" NAME="pfid" VALUE="<% = rsProduct("pfid").Value %>">
                            <INPUT TYPE="SUBMIT" VALUE="Delete Product...">
                        </FORM>
                    </TD>
                    <TD valign=top>
                        <FORM METHOD="POST" ACTION="confirm.asp">
                            <INPUT TYPE="HIDDEN" NAME="confirm_action" VALUE="xt_data_delete.asp">
                            <INPUT TYPE="HIDDEN" NAME="confirm_message" VALUE="Delete all the variants and attributes for '<%= rsProduct("pfid").Value %>'?">
                            <INPUT TYPE="HIDDEN" NAME="type" VALUE="product2">
                            <INPUT TYPE="HIDDEN" NAME="goto" VALUE="product_edit.asp">
                            <INPUT TYPE="HIDDEN" NAME="pfid" VALUE="<% = rsProduct("pfid").Value %>">
                            <INPUT TYPE="SUBMIT" VALUE="Delete Attributes and Variants...">
                        </FORM>
                    </TD>
                </TR>


                <TR>
                    <TD></TD>
                    <TD>
                        <A HREF="<% = "attr_list.asp?" & mscsPage.URLArgs("pfid", Cstr(rsProduct("pfid").Value)) %>">Edit Attribute</A>
                    </TD>
                    <TD>
                        <A HREF="<% = "variant_list.asp?"  & mscsPage.URLArgs("pfid", Cstr(rsProduct("pfid").Value)) %>">Edit Variant</A>
                    </TD>
                </TR>
                            
            </TABLE>
        </TD>
        
        <TD VALIGN=TOP>
            <IMG SRC="/<% = mscsPage.VirtualDirectory() %>/assets/prodimg/<% = rsProduct("image_filename").Value %>">
        </TD>    
    </TR>
</TABLE>

<!--#INCLUDE FILE="include/mgmt_footer.asp" -->
