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
<BODY TOPMARGIN=16 LEFTMARGIN=16 WIDTH=300 BACKGROUND="../assets/images/bnbackground.gif">
</HEAD>
<H2>Error</H2>

<TABLE WIDTH=100% BORDER="0" CELLPADDING="0" CELLSPACING="0">
    <TR>
        <TD>
            We are sorry, but we are unable to process your request:
            <P>
        </TD>
    </TR>
    <TR>
        <TD>
        <UL>
            <%
            Dim errorStr
            for each errorStr in errorList
                Response.Write "<LI>" & errorStr & "</LI>"
            next
            %>
        </UL>
        </TD>
    </TR>
    <TR>
        <TD>
            <br>
            Please go back and correct the error and try again.
        </TD>
    </TR>
    <TR>
        <TD ALIGN="middle">
            <br>
            <FORM>
                <INPUT TYPE=BUTTON   VALUE="Go Back" onClick="history.back()"> 
            </FORM>
        </TD>
    </TR>
</TABLE>

</BODY>
</HTML>



