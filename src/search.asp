<%@ LANGUAGE=vbscript enablesessionstate=false %>
<!--#INCLUDE FILE = "include/shop.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   SEARCH.ASP                                                             #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>

<HTML>

<HEAD>
    <TITLE>Search Page </TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>

<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>" BACKGROUND="<% = background %>" onload="onload()">

<% hd_image = "heading_search.gif"%>
<% alt = "Search Page"%>
<!--#INCLUDE FILE = "include/header.asp" -->

Find products for which the name or description contains the given text phrase:

<P>

<FORM NAME=Search METHOD=POST ACTION="<% = mscsPage.URL("srchresult.asp") %>" >
    Search for: <input type=text size=30 maxlength=256 name="str">
                <input type=button value="Search" OnClick=validate()>
</FORM>

<SCRIPT LANGUAGE="JavaScript">
<!-- 
     function onload(){
        document.Search.str.focus();
     }

     function validate(){
        var String = document.Search.str.value;
        if (String.length > 30){
              alert("Search Item must be 30 characters or less");
              document.Search.str.focus();
              return;
        }
        else
              document.Search.submit();
     }
// -->
</SCRIPT>

<TABLE WIDTH="600">
    <TR>
        <TD>
        <P>
        <I>
           Hint: You can only search for one match at a time. Partial matching works so try to use the root of a word rather than being too explicit. For example, try 'mug' rather than 'mugs' since it will match both.
        </I>
        </TD>
    </TR>
</TABLE>
<% conn.close %>
<!--#INCLUDE FILE = "include/footer.asp" -->

</BODY>
</HTML>
