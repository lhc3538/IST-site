<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
ID = request.QueryString("ID")
LX = request.QueryString("LX")
Operation = request.QueryString("Operation")
Conver = request.QueryString("Conver")
strReferer = Request.ServerVariables("http_referer")

If Operation = "up" Then
    Conn.Execute "update "&LX&" set "&Conver&" = 1 where ID=" & ID
Else
    Conn.Execute "update "&LX&" set "&Conver&" = 0 where ID=" & ID
End If
response.Redirect strReferer
%>