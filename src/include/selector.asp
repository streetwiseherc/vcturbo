<% REM #########################################################################%>
<% REM                                                                          %>
<% REM   SELECTOR.ASP                                                           %>
<% REM   Microsoft Commerce Server v3.00                                        %>
<% REM                                                                          %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.     %>
<% REM                                                                          %>
<% REM #########################################################################%>

<% 
    ' For intranets not connected to the internet, override default
    ' download location here.  For example:
    '  If LCase(CStr(Request("HTTP_UA_CPU"))) <> "alpha" Then
    '     strMSWltIEDwnldLoc  = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & mscspage.siteroot & "/manager/MSCS_Images/controls/MSWallet.cab"
    '  Else
    '     strMSWltIEDwnldLoc  = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & mscspage.siteroot & "/manager/MSCS_Images/controls/MSWltAlp.cab"
    '  End If
    '  strMSWltNavDwnldLoc    = "http://" & Request.ServerVariables("SERVER_NAME") & "/" & mscspage.siteroot & "/manager/MSCS_Images/controls"
%>

<%  ' File Version 3,1,0,1378

    ' Variables set before including this file. 
    '
    ' One of the following two variables should be set before including this file:
    '   fMSWltAddressSelector   Set to "True" when the AddressSelector appears on the
    '                           page.
    '   fMSWltPaymentSelector   Set to "True" when the PaymentSelector appears on the
    '                           page.
    '
    ' The following variables may be optionally set before including this file:
    '   strMSWltIEDwnldLoc      When not connected to the Internet, set to the
    '                           location of the mswallet.cab, the wallet download
    '                           package for IE.
    '   strMSWltNavDwnldLoc     When not connected to the Internet, set to the
    '                           location of the HTM page containing download
    '                           instructions (plginst.htm in the SDK).
    '   strMSWltDwnldVer        Allows overriding the downloaded version of Microsoft
    '                           Wallet.  Overriding the version number should not be
    '                           necessary, so don't use this unless you understand
    '                           what you're doing.
    '   fMSWltShowErrorDialogs  For debugging, set this to "True" to show error
    '                           dialogs in GetValue calls.

    ' Variables set as a side effect of using including this file.  You can use
    ' these in your ASP code:
    '   fMSWltActiveXBrowser        True if the browser supports ActiveX.
    '   fMSWltLiveConnectBrowser    True if the browser DOES NOT support ActiveX and
    '                               DOES support LiveConnect.  This covers Nav3 and
    '                               probably Nav4.
    '   fMSWltUplevelBrowser        True if either of fMSWltActiveXBrowser or
    '                               fMSWltLiveConnectBrowser --- e.g,, it's a control
    '                               case.

    ' Server Side VBScript APIs
    '
    '   MSWltIEAddrSelectorClassid
    '   MSWltIEPaySelectorClassid
    '       Returns the <OBJECT> CLASSID field.  Always use as the CLASSID value.  The
    '       returned classid is different on an Alpha NT machine.
    '
    '   MSWltAddress
    '       Returns HTML text to place in your page. This HTML text is the necessary
    '       text for including the address control.
    '
    '   MSWltIECodebase()
    '       Returns the <OBJECT> CODEBASE field.  Always use as the CODEBASE value.
    '
    '   MSWltNavDwnldURL(strInstructionsFileName)
    '       Returns the <EMBED> PLUGINSPAGE field.  Always use as the PLUGINSPAGE
    '       value.  strInstructionsFileName specifies the name of the instructions
    '       file name.  When using the default download location, call 
    '       MSWltNavDwnldURL with 'plginst.htm'.  When using an Intranet download
    '       location, call MSWltNavDwnldURL with the name of the plugin instructions
    '       file (e.g., plginst.htm in the SDK).
    '       
    '   MSWltLastChanceText(strDownlevelURL, strDownlevelText)
    '       Returns HTML text to place in your page.  This text puts a link reading
    '       "Click here if you have problems with the Microsoft Wallet".  When this
    '       link is clicked, the user navigates to the strDownlevelURL parameter,
    '       which in many cases will be the same page.  This routine automatically
    '       appends "use_form=1" to strDownlevelURL to force the downlevel version of
    '       the page.
    '       We recommend setting the font to size 1.
    '
    '   MSWltLoadDone(strDownlevelURL)
    '       Must be called in the uplevel browser case (fMSWltUplevelBrowser = True) on
    '       the <BODY> OnLoad field. Pass in the downlevel URL, as with 
    '       MSWltLastChanceText.  This is only used when the user refuses the
    '       Nav plugin on initial install or subsequent upgrade.
    '
    '   MSWltAddressControl(fSmall)
    '       Called to include the address control on a page
    '   
    '   MSWltPaymentControl(fSmall, strMSWltAcceptedTypesTypes, strTotal)
    '       Called to include the payment control on a page. The accepttypes and total
    '       are used to control which credit cards are accepted, and the "total" that will
    '       be billed to the card.
    '
    ' Client Side JavaScript APIs.
    '
    '   MSWltLoadDone(strDownlevelURL)
    '       Use the server-side version or strDownlevelURL won't be set correctly.
    '
    '   MSWltPrepareForm(form, cParams, xlationArray)
    '       Use this to update the fields in form with values from the
    '       AddressSelector and/or PaymentSelector.  cParams is the count of
    '       the total number of parameters passed, including the cParams
    '       parameter.  xlationArray is an optional set of parameters used to
    '       translate when form field names to not match field names returned by
    '       the selectors; see the documentation for examples.
    '       Call this routine before posting the form.  This routine will routine
    '       true if is succeeds, so it can be placed in a OnSubmit event handler,
    '       but this is not recommended, because if any JavaScript errors occur, the
    '       post will happen anyway erroneously.
    '       In IE3 when using multiple frames and JavaScript-based navigation 
    '       between frames, using this routine causes subsequent JavaScript-based
    '       navigation to fail by navigating to the wrong location.  We recommend
    '       not using this routine in those cases.  This only applies in the most
    '       complex multi-frame/JavaScript-based navigation cases, like Adventure
    '       Works in CommerceServer v3.00.


    ' Prepare miscellaneous variables.
    Dim fMSWltAddressSelector 
    Dim fMSWltShowErrorDialogs 
    
    fMSWltAddressSelector = CBool(fMSWltAddressSelector)
    fMSWltPaymentSelector = CBool(fMSWltPaymentSelector)
    fMSWltShowErrorDialogs = CBool(fMSWltShowErrorDialogs)

    ' Browser Detection.
    Dim objBrowser
    Dim strCPU
    Dim fMSWltAlphaIE
    Dim fMSWltActiveXBrowser
    Dim fMSWltLiveConnectBrowser
    Dim fMSWltUplevelBrowser
    
    Dim strMSWltDwnldVer
    Dim strMSWltIEDwnldLoc
    Dim strMSWltNavDwnldLoc

    Set objBrowser = Server.CreateObject("MSWC.BrowserType")
    strCPU = LCase(CStr(Request("HTTP_UA_CPU")))    ' CPU is necessary to differentiate between
                                                    ' alpha, x86 and other CPUs on NT.
                                                    ' only set for IE, Nav doesn't set.
    If strCPU = "alpha" Then
        fMSWltAlphaIE = true
    Else
        fMSWltAlphaIE = false
    End If

    ' Note that use_form = 1 forces a downlevel page.

    If Request.QueryString("use_form") = 0 And objBrowser.JavaScript = "True" And _
       (fMSWltAddressSelector Or fMSWltPaymentSelector) And _
       (objBrowser.Platform = "Win95" Or objBrowser.Platform = "Win98" Or _
         ((objBrowser.Platform = "WinNT" Or objBrowser.Platform = "Win32") And _
           ((Len(strCPU) = 0) Or (strCPU = "x86") Or fMSWltAlphaIE))) Then

        If objBrowser.ActiveXControls = "True" Then
            fMSWltActiveXBrowser = True
            fMSWltUplevelBrowser = True
        ElseIf objBrowser.Browser = "Netscape" And _
               ( (CInt(objBrowser.majorver) > 3) Or _
                 ((CInt(objBrowser.majorver) = 3) And (objBrowser.beta = "False")) _
               ) Then
            fMSWltLiveConnectBrowser = True
            fMSWltUplevelBrowser = True
        Else
            fMSWltActiveXBrowser = False
            fMSWltLiveConnectBrowser = False
            fMSWltUplevelBrowser = False
        End If
    Else
        fMSWltActiveXBrowser = False
        fMSWltLiveConnectBrowser = False
        fMSWltUplevelBrowser = False
    End If


    ' Examine to see if the download version or location has
    ' been overridden.  These should be overridden only when the 
    ' consumer is not connected to the Internet (i.e, intranet scenario).
    ' When not overriden, the default locations and version on
    ' Microsoft.com will be used.

    ' Download version.

    If IsEmpty(strMSWltDwnldVer) Then
        strMSWltDwnldVer = "2,1,0,1378"
    Else
        strMSWltDwnldVer = CStr(strMSWltDwnldVer)
    End If

    ' IE Wallet version location.
    If IsEmpty(strMSWltIEDwnldLoc) Then
        If fMSWltAlphaIE Then
            strMSWltIEDwnldLoc = "mswltalp.cab"
        Else
            strMSWltIEDwnldLoc = "mswallet.cab"
        End If
    Else
        strMSWltIEDwnldLoc = CStr(strMSWltIEDwnldLoc)
    End If

    ' Navigator Wallet version location.

    If IsEmpty(strMSWltNavDwnldLoc) Then
        strMSWltNavDwnldLoc = "http://www.microsoft.com/commerce/wallet/local/"
    Else
        strMSWltNavDwnldLoc = CStr(strMSWltNavDwnldLoc)
        ' add trailing blank if missing
        If Right(strMSWltNavDwnldLoc, 1) <> "/" Then
            strMSWltNavDwnldLoc = strMSWltNavDwnldLoc & "/"
        End If
    End If

    Function MSWltIEAddrSelectorClassid
        If fMSWltAlphaIE Then
            MSWltIEAddrSelectorClassid = "clsid:B7FB4D5B-9FBE-11d0-8965-0000F822DEA9"
        Else
            MSWltIEAddrSelectorClassid = "clsid:87D3CB63-BA2E-11cf-B9D6-00A0C9083362"
        End If
    End Function

    Function MSWltIEPaySelectorClassid
        If fMSWltAlphaIE Then
            MSWltIEPaySelectorClassid = "clsid:B7FB4D5C-9FBE-11d0-8965-0000F822DEA9"
        Else
            MSWltIEPaySelectorClassid = "clsid:87D3CB66-BA2E-11cf-B9D6-00A0C9083362"
        End If
    End Function

    Function MSWltIECodebase
        MSWltIECodebase = strMSWltIEDwnldLoc & "#Version=" & strMSWltDwnldVer
    End Function

    Function MSWltNavDwnldURL(strInstructionsFileName)
        MSWltNavDwnldURL = strMSWltNavDwnldLoc & CStr(strInstructionsFileName)
    End Function

    ' Tack use_form=1 on to the end of a URL
    Function MSWltDownlevelURL(strDownlevelURL)
        Dim nQmarkLoc
        MSWltDownlevelURL = Trim(CStr(strDownlevelURL)) ' any spaces lurking?  get rid of them
        nQmarkLoc = InStr(MSWltDownlevelURL, "?")
        If nQmarkLoc > 0 Then
            If nQmarkLoc = Len(MSWltDownlevelURL) Then
                MSWltDownlevelURL = MSWltDownlevelURL & "use_form=1"
            Else
                MSWltDownlevelURL = MSWltDownlevelURL & "&use_form=1"
            End If
        Else
            MSWltDownlevelURL = MSWltDownlevelURL & "?use_form=1"
        End If
        MSWltDownlevelURL = MSWltDownlevelURL
    End Function

    Function MSWltLastChanceText(strDownlevelURL, strDownlevelText)
        If fMSWltUplevelBrowser And (fMSWltAddressSelector Or fMSWltPaymentSelector) Then
            MSWltLastChanceText = " <A HREF=""" & _
                                  MSWltDownlevelURL(strDownlevelURL) & _
                                  """ > " & strDownlevelText & "</A> "
        End If
    End Function

    Function MSWltLoadDone(strDownlevelURL)
        ' Call JavaScript MSWltLoadDone routine with downlevel URL
        MSWltLoadDone = "MSWltLoadDone('" & MSWltDownlevelURL(strDownlevelURL) & "')"
    End Function

    function MSWltPaymentControl(small, strMSWltAcceptedTypesTypes, strTotal)
        Dim width
        Dim height
        Dim useComboBox
        if small = True then
            width = 200
            height = 30
            useComboBox = "true"
        else
            width = 154
            height = 123
            useComboBox = "false"
        end if

        if fMSWltActiveXBrowser Then  %>
            <OBJECT
                            ID=paySelector
                            CLASSID="<% = MSWltIEPaySelectorClassid() %>"
                            CODEBASE="<% = MSWltIECodebase() %>"
                            HEIGHT="<% = height %>"
                            WIDTH="<% = width %>">
                            <PARAM NAME="UseComboBox" VALUE="<% = useComboBox %>">                            
                            <PARAM NAME="AcceptedTypes" VALUE="<% = strMSWltAcceptedTypes %>">
                            <PARAM NAME="Total" VALUE="<% = strTotal %>">
            </OBJECT>
        <% elseif fMSWltLiveConnectBrowser then %>
            <EMBED
                             NAME="paySelector"
                             TYPE="application/x-mswallet"
                             PLUGINSPAGE="<% = MSWltNavDwnldURL("plginst.htm") %>"
                             VERSION="<% = strMSWltDwnldVer %>" 
                             HEIGHT="<% = Height %>"
                             WIDTH="<% = Width %>"
                             ACCEPTEDTYPES="<% = strMSWltAcceptedTypes %>"
                             TOTAL="<% = strTotal %>"
                             USECOMBOBOX="<% = useComboBox %>">
        <% end if
    end function

    function MSWltAddressControl(small)
        Dim width
        Dim height
        Dim useComboBox
        if small = True then
            width = 200
            height = 30
            useComboBox = "true"
        else
            width = 154
            height = 123
            useComboBox = "false"
        end if
    
        if fMSWltActiveXBrowser then %>  
            <OBJECT ID=addrSelector
                                CLASSID="<% = MSWltIEAddrSelectorClassid() %>"
                                CODEBASE="<% = MSWltIECodebase() %>"
                                HEIGHT=<% = height %>
                                WIDTH=<% = width %>>
                                <PARAM NAME="UseComboBox" VALUE="<% = useComboBox %>">
            </OBJECT>
        <% elseif fMSWltLiveConnectBrowser then %>
            <EMBED
                            NAME="addrSelector"
        		    		TYPE="application/x-msaddr"
                            PLUGINSPAGE="<%= MSWltNavDwnldURL("plginst.htm") %>"
                            VERSION="<%= strMSWltDwnldVer %>"
                            HEIGHT="<% = height %>"
                            WIDTH="<% = width %>"
                            USECOMBOBOX="<% = useComboBox %>">
        <% end if
    end function 
%>

<% ' JavaScript to insert to support the controls. %>
<% If fMSWltUplevelBrowser Then %>
<SCRIPT LANGUAGE="JavaScript">

    var fMSWltLoaded = false    <% ' Has onLoad initialization been done? %>

    <% If fMSWltAddressSelector Then %>
        var objAddrSelector     <% ' Address selector from both IE and Nav %>
    <% End If %>

    <% If fMSWltPaymentSelector Then %>
        var objPaySelector      <% ' Payment selector from both IE and Nav %>
    <% End If %>

    // JavaScript version.
    function MSWltLoadDone(strDownlevelURL)
    {
        <% If fMSWltLiveConnectBrowser Then %>
            <% ' Is the plugin around? %>

            if (
                <% If fMSWltAddressSelector Then %>
                    (document.addrSelector == null)
                    <% If fMSWltPaymentSelector Then %>
                        ||
                    <% End If %>
                <% End If %>
                <% If fMSWltPaymentSelector Then %>
                    (document.paySelector == null)
                <% End If %>
                )
            {
                if (confirm("Click OK to install Microsoft Wallet Plugins."))
                    <% ' open instructions page in a new window %>
                    window.open("<% = MSWltNavDwnldURL("plginst.htm") %>")
                else
                    location = strDownlevelURL
                return
            }

            <% ' Take care of naming differences between objects and plugins. %>
            <% If fMSWltAddressSelector Then %>
                objAddrSelector = document.addrSelector
                fVersionOK = objAddrSelector.VersionCheck()
            <% End If %>

            <% If fMSWltAddressSelector And fMSWltPaymentSelector Then %>
                if (fVersionOK)
                {
                    objPaySelector = document.paySelector
                    fVersionOK = objPaySelector.VersionCheck()
                }
            <% ElseIf fMSWltPaymentSelector Then %>
                objPaySelector = document.paySelector
                fVersionOK = objPaySelector.VersionCheck()
            <% End If %>

            <% ' Check plugin version.  Version requested set on <embed> tag. %>
            <% ' This version is checked against the version resource in the DLL. %>
            if (!fVersionOK)
            {
                if (confirm("Your Microsoft Wallet Plugins are out of date and you need to upgrade.\nClick OK to view upgrade directions."))
                    <% ' open instructions page in a new window %>
                    window.open("<% = MSWltNavDwnldURL("plginst.htm") %>")
                else
                    location = strDownlevelURL
                return
            }
            fMSWltLoaded = true
        <% End If %>
    }

    function MSWltCheckLoaded()
    {
        if (!fMSWltLoaded)
        {
            <% If fMSWltActiveXBrowser Then %>

                if (
                    <% If fMSWltAddressSelector Then %>
                        (!(!addrSelector))
                        <% If fMSWltPaymentSelector Then %>
                            && 
                        <% End If %>
                    <% End If %>
                    <% If fMSWltPaymentSelector Then %>
                        (!(!paySelector))
                    <% End If %>
                    )
                {
                    <% If fMSWltAddressSelector Then %>
                        objAddrSelector = addrSelector
                    <% End If %>
                    <% If fMSWltPaymentSelector Then %>
                        objPaySelector = paySelector
                    <% End If %>
                    fMSWltLoaded = true
                }
                else
            <% End If %>
            
                    <% ' Navigator JavaScript does not take kindly to pushing buttons before the page is done %>
                    <% ' loading; hence the warning that a reload and wait may be necessary. %>
                    alert("Page not done loading yet.  Try again when it's loaded.  Refresh the page if you're having difficulty (then wait for it to load).")
        }
        return fMSWltLoaded
    }

    function doNothing() { <% ' Do nothing, supports object creation with no contents. %> }

    function MSWltPrepareForm(form, cParams, xlationArray)
    {   
        if (!MSWltCheckLoaded())
            return false

        <% If fMSWltPaymentSelector Then %>
            PI = objPaySelector.GetValues()           <% ' get payment information (PI) %>
            errorStatus = objPaySelector.GetLastError()  <% ' did an error occur?        %>
            if (errorStatus < 0)
            {
                if (errorStatus != (-2147220991) && errorStatus != (-2147220990))  <% ' HRESULT 0x80040201, WALLET_E_CANCEL, and HRESULT 0x80040202, WALLET_E_HANDLEDERROR %>
                    alert("Payment selection failed due to an unknown problem.")
                return false
            }
        <% End If %>

        <% If fMSWltAddressSelector Then %>
            shipTo = objAddrSelector.GetValues()           <% ' get ship to address information %>

            errorStatus = objAddrSelector.GetLastError()  <% ' did an error occur? %>
            if (errorStatus < 0)
            {
                if (errorStatus != (-2147220991) && errorStatus != (-2147220990))  <% ' HRESULT 0x80040201, WALLET_E_CANCEL, and HRESULT 0x80040202, WALLET_E_HANDLEDERROR %>
                    alert("Address selection failed due to an unknown problem.")
                return false
            }
        <% End If %>

        elements = form.elements


        <% ' Build xlation table from xlationArray %>
        xlate = new doNothing()

        for (i = 2; i < cParams; i += 2)
        {
            value = MSWltPrepareForm.arguments[i+1]
            if (value.length > 0)
                xlate[MSWltPrepareForm.arguments[i]] = value
        }

        <% If fMSWltPaymentSelector Or fMSWltAddressSelector Then %>
            for (i = 0; i < elements.length; i++)
            { 
                if (form.elements[i].name.length > 0)
                {
                    xlateValue = xlate[elements[i].name]
                    if (xlateValue)
                        name = xlateValue
                    else
                        name = elements[i].name

                    <% If fMSWltPaymentSelector Then %>
                        <% ' Have to make the string at least 1 long to get around Nav issue %>
                        value = 'a' + objPaySelector.GetValue(PI, name, <% = -CInt(fMSWltShowErrorDialogs) %>)
                        if (value.length > 1)
                            elements[i].value = value.substring(1)
                    <% End If %>

                    <% If fMSWltAddressSelector Then %>
                        <% ' Have to make the string at least 1 long to get around Nav issue %>
                        value = 'a' + objAddrSelector.GetValue(shipTo, name, <% = -CInt(fMSWltShowErrorDialogs) %>)
                        if (value.length > 1)
                            elements[i].value = value.substring(1)
                    <% End If %>
                }
            }
        <% End If %>

        return true
    }

</SCRIPT>
<% End If %>
