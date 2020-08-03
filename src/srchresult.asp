<%@ LANGUAGE=vbscript  %>
<!--#INCLUDE FILE = "include/shop.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   SEARCH_RESULT.ASP                                                      #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>

<HTML>

<HEAD>
    <TITLE>Search Result Page </TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>

<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>" BACKGROUND="<% = background %>">

<% hd_image = "heading_search.gif" %>
<% alt = "Search Results Page"%>
<!--#INCLUDE FILE = "include/header.asp" -->

<H1>Search Results</H1>

<% if Request("str") <> "" then
      Dim strText
      Dim rsSrchproducts
      strText = "'%" & Replace(Request("str"),"'","''") & "%'"
      set rsSrchproducts = Conn.Execute("select pfid, dept_id, name from vcturbo_product_family  where  name like " & strText & "  or  short_description like " & strText & " order by dept_id")

      if rsSrchproducts.EOF then %>
            No products found for which name or description contains '<% = mscsPage.HTMLEncode(Request("str")) %>'.
      <% else %>
            Below are all products for which name or description contains '<% = mscsPage.HTMLEncode(Request("str")) %>':
            <P>
            <% While not rsSrchproducts.EOF %>
                <A HREF = "<% = mscsPage.URL("product.asp" , "pfid", rsSrchproducts("pfid").Value) %>"><% = mscsPage.HTMLEncode(rsSrchproducts("name").Value) %></A><br>
                <% rsSrchproducts.Movenext
             Wend %>
        <% end if %>
<% else %>
     No products found for which name or description contains '<% = mscsPage.HTMLEncode(Request("str")) %>'.
<% end if %>
<% conn.close %>
<!--#INCLUDE FILE = "include/footer.asp" -->

</BODY>
</HTML>
