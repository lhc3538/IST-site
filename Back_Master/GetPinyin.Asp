<!--#include file="../Include/Function.Asp" -->
<%
SortName = request("SortName")
response.Write Pinyin(trim(SortName))
%>