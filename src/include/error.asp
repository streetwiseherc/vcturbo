<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   ERROR.ASP                                                              #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>


<HTML>
<HEAD>
    <TITLE>Error</TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>

<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>" background="<% = background %>" topmargin=20 leftmargin=20>

<% alt="Error" %>
<!--#INCLUDE FILE="header.asp" -->
        We are sorry, but we are unable to process your request:
        <BR>
        <UL>
        <% Dim errorStr
        for each errorStr in errorList
             Response.Write "<LI>" & errorStr & "</LI>"
        next
        %>
        </UL>
        Please go back and correct the error and try again.
        <BR>
        <FORM>
            <INPUT TYPE=BUTTON   VALUE="Go Back" onClick="history.back()"> 
        </FORM>
<!--#INCLUDE FILE = "footer.asp" -->

</BODY>
</HTML>
