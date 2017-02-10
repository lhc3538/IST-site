<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
Dim rs, sql, ID, SortName
ID = request.QueryString("ID")
SortName = request.QueryString("SortName")
set rs = server.CreateObject("adodb.recordset")
rs.open "select * from keyicms_Template where UseFlag and SortName='"&SortName&"'",conn,1,2
do until rs.eof
	rs("UseFlag") = 0
	rs.update
	rs.movenext
loop
rs.close

rs.open "select * from keyicms_Template where ID = "&ID,conn,1,2
rs("UseFlag") = -1
rs.update
rs.close
set rs = nothing

Response.redirect request.servervariables("http_referer")
%>