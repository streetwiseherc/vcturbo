<% Option Explicit %>
<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   SHOP.ASP                                                                %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<% REM This file gets included by all Commerce Server pages %>

<% 
function this_page(nowat) 
    this_page = ( Right(Request.ServerVariables("SCRIPT_NAME"),len(nowat)) = nowat)
end function
%>


<OBJECT RUNAT=Server ID=conn     PROGID="ADODB.Connection"></OBJECT>
<OBJECT RUNAT=Server ID=mscsPage PROGID="Commerce.Page">   </OBJECT> 

<%
Dim hd_image
Dim alt
Dim cmdTemp

REM -- ADO command types
Dim adCmdText
Dim adCmdTable
Dim adCmdStoredProc
Dim adCmdUnknown

adCmdText       = 1 
adCmdTable      = 2 
adCmdStoredProc = 4
adCmdUnknown    = 8

REM -- ADO cursor types
Dim adOpenForwardOnly
Dim adOpenKeyset
Dim adOpenDynamic 
Dim adOpenStatic 

adOpenForwardOnly = 0
adOpenKeyset      = 1
adOpenDynamic     = 2
adOpenStatic      = 3

REM -- ADO lock types
Dim adLockReadOnly        
Dim adLockPessimistic     
Dim adLockOptimistic      
Dim adLockBatchOptimistic 

adLockReadOnly        = 1
adLockPessimistic     = 2
adLockOptimistic      = 3
adLockBatchOptimistic = 4

REM page constants
Dim color_text  
Dim color_link  
Dim color_alink 
Dim color_vlink 
Dim background 

color_text  = "#51392B"
color_link  = "#cc9966"
color_alink = "#00FF00"
color_vlink = "#FF0000"
background = "assets/images/main.gif"


REM -- If store is not open then redirect to closed URL
if MSCSSite.Status <> "Open" then
    response.redirect(MSCSSite.CloseRedirectURL)
end if

REM ******************************************************
REM -- functions for faster page links

function pageURL(pageName)
    pageURL = rootURL & pageName & "?" & emptyArgs
end function

function pageSURL(pageName)
    pageSURL = rootSURL & pageName & "?" & emptyArgs
end function

function baseURL(pageName)
    REM -- you must put on your own shopperArgs
    baseURL = rootURL & pageName & "?"
end function

function baseSURL(pageName)
    REM -- you must put on your own shopperArgs
    baseSURL = rootSURL & pageName & "?"
end function

Dim displayName 
Dim siteRoot    
Dim rootURL     
Dim rootSURL    
Dim emptyArgs   

displayName = mscsPage.HTMLEncode(MSCSSite.DisplayName) 
siteRoot    = mscsPage.SiteRoot()
rootURL     = mscsPage.URLPrefix() & "/" & siteRoot & "/"
rootSURL    = mscsPage.SURLPrefix() & "/" & siteRoot & "/"
emptyArgs   = mscsPage.URLShopperArgs()

REM ******************************************************

REM -- Manually create shopper id
Dim mscsShopperID
mscsShopperID = mscsPage.GetShopperId
if IsNull(mscsShopperID) then
    if Not this_page("default.asp") and Not this_page("welcome_lookup.asp") and Not this_page("welcome_new.asp") and not this_page("xt_shopper_lookup.asp") and not this_page("xt_shopper_new.asp") then
        Response.Redirect(mscsPage.URL("default.asp"))          
    end if
end if

REM -- Create ADO Connection and Command Objects
conn.Open MSCSSite.DefaultConnectionString
%>


