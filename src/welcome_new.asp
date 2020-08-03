<%@ LANGUAGE=vbscript%>
<!--#INCLUDE FILE = "include/shop.asp" -->
<% REM ##########################################################################%>
<% REM                                                                          #%>
<% REM   WELCOME_NEW.ASP                                                        #%>
<% REM   Microsoft Commerce Server v3.00                                        #%>
<% REM                                                                          #%>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     #%>
<% REM                                                                          #%>
<% REM ##########################################################################%>


<HTML>

<HEAD>
    <TITLE>New Shopper</TITLE>
    <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
</HEAD>

<BODY TEXT="<% = color_text %>" LINK="<% = color_link %>" VLINK="<% = color_vlink %>" ALINK="<% = color_alink %>" BACKGROUND="<% = background %>">

<% hd_image="heading_welcome.gif" %>
<% alt="Welcome" %>
<!--#INCLUDE FILE = "include/header.asp" -->

<%' Text Table %>
<TABLE WIDTH=600>
    <TR>
        <TD>                
        Welcome to the <% = displayName %> sign-up sheet.
        To make it easier to participate in the exciting <% = displayName %> community we ask that you take
        a moment to tell us a bit about yourself.
        <P>
        We'll keep all this information to ourselves (naturally); it just helps us serve you better. We'll know how to keep you 
        informed about our events and programs and we'll be able to keep track of what
        you order and where to ship it.
        <P>
        Just fill out the form below and click the "Continue" button.
        </TD>
    </TR>
</TABLE>

<FORM METHOD=POST ACTION="<% = pageSURL("xt_shopper_new.asp") %>" NAME="NewUser">
    <INPUT TYPE="hidden" SIZE=10 NAME="country" VALUE="USA">
    
<%' Registration Table %>
    <TABLE BORDER="0" WIDTH=600>
        <TR>
            <TD ALIGN="CENTER">
                <TABLE CELLSPACING="1" CELLPADDING="1" BORDER="0">
                     <TR><TD>Name: </TD>
                         <TD colspan=3><INPUT TYPE="TEXT" SIZE=25 MAXLENGTH=235 NAME="name"></TD>
                         <TD WIDTH=8>&nbsp;</TD>
                     </TR>
                     <TR><TD> Street: </TD>
                         <TD colspan=3><INPUT TYPE="TEXT" SIZE=25 MAXLENGTH=50 NAME="street"></TD>
                         <TD WIDTH=8>&nbsp;</TD>
                     </TR>
                     <TR><TD> City: </TD>
                         <TD colspan=3><INPUT TYPE="TEXT" SIZE=25 MAXLENGTH=50 NAME="city"></TD>
                         <TD WIDTH=0>&nbsp;</TD>
                     </TR>
                     <TR><TD> State: </TD>
                         <TD><INPUT TYPE="TEXT" SIZE=2 MAXLENGTH=30 NAME="state"></TD>
                         <TD> Zip Code:  </TD>
                         <TD ALIGN="LEFT"> <INPUT TYPE="TEXT" SIZE=10 MAXLENGTH=15 NAME="zip"></TD>
                         <TD WIDTH=0>&nbsp;</TD>
                     </TR>
                </TABLE>
            </TD>
            <TD VALIGN="top" ALIGN="LEFT">
                <TABLE BORDER="0">
                    <TR><TD>Phone: </TD>
                        <TD><INPUT TYPE="TEXT" SIZE=25 MAXLENGTH=16 NAME="phone"></TD>
                        <TD>&nbsp;</TD>
                    </TR>
                    <TR><TD>E-mail: </TD>
                        <TD><INPUT TYPE="TEXT" SIZE=25 MAXLENGTH=50 NAME="email"></TD>
                        <TD>&nbsp;</TD>
                    </TR>
                </TABLE>
            </TD>
       </TR>
       <TR>
           <TD colspan=6>Create and then confirm the password you will use when you visit this store again in the future.<BR> This password must be at least 4 characters long:</td>
       </TR>
       <TR>
           <TD COLSPAN=6 ALIGN=LEFT>
               <TABLE BORDER="0">
                   <TR><TD>Password:</TD>
                       <TD colspan=0> <INPUT TYPE=password SIZE=25 MAXLENGTH=20 NAME="pwd1"></TD>
                       <TD WIDTH=14>&nbsp;</TD> 
                       <TD>Confirm<BR>Password:</TD>
                       <TD colspan=0><INPUT TYPE=password SIZE=25 MAXLENGTH=20 NAME="pwd2" ></TD>
                   <TR>
               </TABLE> 
           </TD>
        </TR>
        <TR>
            <TD WIDTH="530" ALIGN=RIGHT COLSPAN="2"><input type=image SRC="assets/images/btn_continue.gif" ALIGN=BOTTOM BORDER=0>&nbsp;</TD>
       </TR>
    </TABLE>

</FORM>
<% conn.close %>
<!--#INCLUDE FILE = "include/footer.asp" -->

</BODY>
</HTML>

