<% REM ######################################################################### %>
<% REM                                                                           %>
<% REM   LIST.ASP                                                                %>
<% REM   Microsoft Commerce Server v3.00                                         %>
<% REM                                                                           %>
<% REM   Copyright (c) 1996-98 Microsoft Corporation.  All rights reserved.      %>
<% REM                                                                           %>
<% REM ######################################################################### %>

<%

REM -- initialize page variables passed in
if IsNull(ListPageSize) or IsEmpty(ListPageSize) then
	ListPageSize = 5
end if

if IsNull(ListNoRowsText) or IsEmpty(ListNoRowsText) then
	ListNoRowsText = "No data available"
end if

REM -- get control for going from page to page (from form)
PagingMove = mscsPage.RequestString("list_PagingMove", "")

REM Get the page we are on
ListAbsolutePage = mscsPage.RequestNumber("ListAbsolutePage", 1)

REM -- load recordset
cmdTemp.CommandText = ListSQL
Set rsList = Server.CreateObject("ADODB.Recordset")
rsList.CacheSize = ListPageSize
rsList.Open cmdTemp, , adOpenStatic, adLockReadOnly

REM check for empty record set (beginning and end the same or can't get EOF or BOF)
EmptyRecordset = False
On Error Resume Next
If rsList.BOF And rsList.EOF Then EmptyRecordset = True
On Error Goto 0
If Err Then EmptyRecordset = True


REM Set the page size in the record set
rsList.PageSize = ListPageSize


REM -- act on pagemove instruction (from form)
If Not EmptyRecordset Then
    If Not rsList.Supports(adApproxPosition) then 'not SQL Server - probably Oracle
        While Not rsList.EOF
            RecordCount = RecordCount + 1
            rsList.MoveNext
        Wend
        rsList.MoveFirst
        PageCount = Round(RecordCount/ListPageSize)
        If PageCount < RecordCount/ListPageSize then PageCount = PageCount + 1
    else
        PageCount   = rsList.PageCount
        RecordCount = rsList.RecordCount
    end if
    Select Case PagingMove
        Case "<<"
            ListAbsolutePage = 1
        Case "<"
            ListAbsolutePage = ListAbsolutePage - 1
        Case ">"
            ListAbsolutePage = ListAbsolutePage + 1
        Case ">>"
            ListAbsolutePage = PageCount
        Case Else
    End Select
    If (ListAbsolutePage - 1) * ListPageSize > RecordCount or ListAbsolutePage < 1 Then ListAbsolutePage = 1
    If Not rsList.Supports(adApproxPosition) then 'not SQL Server - probably Oracle
        StartRecord = ((ListAbsolutePage - 1) * ListPageSize)
        For i = 1 to StartRecord
            If rsList.EOF then exit for
            rsList.MoveNext
        Next
    else
        rsList.AbsolutePage = ListAbsolutePage
    end if
    If rsList.EOF Then
        ListAbsolutePage = ListAbsolutePage - 1
        If Not rsList.Supports(adApproxPosition) then 'not SQL Server - probably Oracle
            rsList.MoveFirst
            StartRecord = ((ListAbsolutePage - 1) * ListPageSize)
            For i = 1 to StartRecord
                If rsList.EOF then exit for
                rsList.MoveNext
            Next
        else
            rsList.AbsolutePage = ListAbsolutePage
        end if
    End If
End If

REM -- show navigation bar (if more than 1 page)
If PageCount > 1 and Not EmptyRecordset Then
%>
        <TABLE BORDER=0 CELLSPACING=3>
        <TR>
            <TH ALIGN=CENTER BGCOLOR=lightgrey>
                <NOBR><% = RecordCount %> Records &nbsp;&nbsp;&nbsp;&nbsp; Page: <% = ListAbsolutePage %> of <% = PageCount %></NOBR>
            </TH>
        </TR>
        <TR>
            <TD>
                <FORM ACTION="<% = Request.ServerVariables("PATH_INFO") %>" METHOD=POST>
<% If ListAbsolutePage > 1 Then %>
                    <INPUT TYPE="Submit" NAME="list_PagingMove" VALUE="   &lt;&lt;   ">
                    <INPUT TYPE="Submit" NAME="list_PagingMove" VALUE="   &lt;    ">
<% Else %>
                    <INPUT TYPE="Button" NAME="list_PagingMove" VALUE="   &lt;&lt;   ">
                    <INPUT TYPE="Button" NAME="list_PagingMove" VALUE="   &lt;    ">
<% End If %>
<% If ListAbsolutePage < PageCount Then %>
                    <INPUT TYPE="Submit" NAME="list_PagingMove" VALUE="    &gt;   ">
                    <INPUT TYPE="Submit" NAME="list_PagingMove" VALUE="   &gt;&gt;   ">
<% Else %>
                    <INPUT TYPE="Button" NAME="list_PagingMove" VALUE="    &gt;   ">
                    <INPUT TYPE="Button" NAME="list_PagingMove" VALUE="   &gt;&gt;   ">
<% End If %>
                    <INPUT TYPE="Submit" NAME="list_PagingMove" VALUE=" Requery ">
                    <INPUT TYPE="Hidden" NAME="ListAbsolutePage" VALUE="<% = ListAbsolutePage %>">
                </FORM>
            </TD>
        </TR>
        </TABLE>
<%
End If
%>


<%
REM -- show no rows message if the recordset is empty
If EmptyRecordset Then
	Response.Write("<I>" & ListNoRowsText & "</I>")
REM -- else show the row information for each row
Else
%>
<TABLE BORDER=0 CELLPADDING=2>
    <TR> <% call ListHeader() %></TR>
<%
    RowCount = ((ListAbsolutePage - 1) * ListPageSize) + 1
	RecordsProcessed = 0
    while not rsList.EOF and RecordsProcessed < ListPageSize
        RecordsProcessed = RecordsProcessed + 1
        Response.Write("<TR>")
        call ListRow()
        Response.Write("<TR>")
        rsList.MoveNext
        RowCount = RowCount + 1
	wend
%>
</TABLE>
<%
End If
%>
<P>
