<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/System_Language.asp" -->
<!--#include file="../Include/Version.Asp" -->
<!--#include file="html_Config.Asp"-->
<!--#include file="html_Function.Asp"-->

<!--#include file="html_To_About.Asp"-->
<!--#include file="html_To_AboutSort.Asp"-->
<!--#include file="html_To_Case.Asp"-->
<!--#include file="html_To_CaseSort.Asp"-->
<!--#include file="html_To_Download.Asp"-->
<!--#include file="html_To_DownSort.Asp"-->
<!--#include file="html_To_Job.Asp"-->
<!--#include file="html_To_JobSort.Asp"-->
<!--#include file="html_To_News.Asp"-->
<!--#include file="html_To_NewsSort.Asp"-->
<!--#include file="html_To_Product.Asp"-->
<!--#include file="html_To_ProductSort.Asp"-->
<!--#include file="html_To_Index.Asp"-->

<%
'生成静态AJAX调用
Call SiteInfo()
Table = request("Table")
ID = request("ID")
FileName = request("FileName")
Ke01 = request("Ke01")
If Ke01 = "PC" Then Flag = " ViewFlag " : TemplatePath = TemplatePath
If Ke01 = "Mobile" Then Flag = " MobileFlag " : TemplatePath = mTemplatePath

Select Case Table
	Case "About"
		Call html_To_About(ID, Ke01)
	Case "AboutSort"
		Call html_To_AboutSort("Index."&HTMLName, Ke01)
	Case "Case"
		Call html_To_Case(ID, Ke01)
	Case "CaseSort"
		set strDictionary = server.CreateObject("SCRIPTING.DICTIONARY")
		Set regExpObj = New Regexp
		regExpObj.IgnoreCase = True 
		regExpObj.Global = True
		regExpObj.Pattern="{keyicms:CaseList([\s\S]*?)}([\s\S]*?){/keyicms:CaseList}"
		replace_code = ReadFromUTF(TemplatePath&"/html/CaseList.html","utf-8")
		set matches = regExpObj.Execute(replace_code)
		for each match in matches
			labelStr = match.SubMatches(0)
			set labelArr = parseArr(labelStr)
			lNum = labelArr("Num")
		next
		set matches = nothing
		Set regExpObj = nothing
		set strDictionary = nothing

		If ID <> "" Then
			SortPath = Conn.execute("Select * from keyicms_CaseSort where ID="&ID)("SortPath")
			datawhere =  " and instr(SortPath,'"&SortPath&"')>0"
		End If
		
		set rs = server.CreateObject("adodb.recordset")
		rs.open "Select * from keyicms_Case Where " & Flag & datawhere, conn, 1, 2
		rs.pagesize = lNum
		totalpage = rs.pagecount
		rs.close : set rs = nothing
		If totalpage = 0 then totalpage = totalpage + 1
		
		Call html_To_CaseSort(ID, "Index."&HTMLName, "", Ke01)
		For n = 1 To totalpage
			Call html_To_CaseSort(ID, "Index"&Separated&n&"."&HTMLName, n, Ke01)
		Next

	Case "Download"
		Call html_To_Download(ID, Ke01)
	Case "DownSort"
		set strDictionary = server.CreateObject("SCRIPTING.DICTIONARY")
		Set regExpObj = New Regexp
		regExpObj.IgnoreCase = True 
		regExpObj.Global = True
		regExpObj.Pattern="{keyicms:DownList([\s\S]*?)}([\s\S]*?){/keyicms:DownList}"
		replace_code = ReadFromUTF(TemplatePath&"/html/DownList.html","utf-8")
		set matches = regExpObj.Execute(replace_code)
		for each match in matches
			labelStr = match.SubMatches(0)
			set labelArr = parseArr(labelStr)
			lNum = labelArr("Num")
		next
		set matches = nothing
		Set regExpObj = nothing
		set strDictionary = nothing

		If ID <> "" Then
			SortPath = Conn.execute("Select * from keyicms_DownSort where ID="&ID)("SortPath")
			datawhere =  " and instr(SortPath,'"&SortPath&"')>0"
		End If
		
		set rs = server.CreateObject("adodb.recordset")
		rs.open "Select * from keyicms_Download Where "&Flag&datawhere, conn, 1, 2
		rs.pagesize = lNum
		totalpage = rs.pagecount
		rs.close : set rs = nothing
		If totalpage = 0 then totalpage = totalpage + 1
		
		Call html_To_DownSort(ID, "Index."&HTMLName, "", Ke01)
		For n = 1 To totalpage
			Call html_To_DownSort(ID, "Index"&Separated&n&"."&HTMLName, n, Ke01)
		Next

	Case "Job"
		Call html_To_Job(ID, Ke01)
	Case "JobSort"
		set strDictionary = server.CreateObject("SCRIPTING.DICTIONARY")
		Set regExpObj = New Regexp
		regExpObj.IgnoreCase = True 
		regExpObj.Global = True
		regExpObj.Pattern="{keyicms:JobList([\s\S]*?)}([\s\S]*?){/keyicms:JobList}"
		replace_code = ReadFromUTF(TemplatePath&"/html/JobList.html","utf-8")
		set matches = regExpObj.Execute(replace_code)
		for each match in matches
			labelStr = match.SubMatches(0)
			set labelArr = parseArr(labelStr)
			lNum = labelArr("Num")
		next
		set matches = nothing
		Set regExpObj = nothing
		set strDictionary = nothing

		set rs = server.CreateObject("adodb.recordset")
		rs.open "Select * from keyicms_Job Where "&Flag&datawhere, conn, 1, 2
		rs.pagesize = lNum
		totalpage = rs.pagecount
		rs.close : set rs = nothing
		If totalpage = 0 then totalpage = totalpage + 1
		
		Call html_To_JobSort("Index."&HTMLName, "", Ke01)
		For n = 1 To totalpage
			Call html_To_JobSort("Index"&Separated&n&"."&HTMLName, n, Ke01)
		Next

	Case "News"
		Call html_To_News(ID, Ke01)
	Case "NewsSort"
		set strDictionary = server.CreateObject("SCRIPTING.DICTIONARY")
		Set regExpObj = New Regexp
		regExpObj.IgnoreCase = True 
		regExpObj.Global = True
		regExpObj.Pattern="{keyicms:NewsList([\s\S]*?)}([\s\S]*?){/keyicms:NewsList}"
		replace_code = ReadFromUTF(TemplatePath&"/html/NewsList.html","utf-8")
		set matches = regExpObj.Execute(replace_code)
		for each match in matches
			labelStr = match.SubMatches(0)
			set labelArr = parseArr(labelStr)
			lNum = labelArr("Num")
		next
		set matches = nothing
		Set regExpObj = nothing
		set strDictionary = nothing

		If ID <> "" Then
			SortPath = Conn.execute("Select * from keyicms_NewsSort where ID="&ID)("SortPath")
			datawhere = " and instr(SortPath,'"&SortPath&"')>0"
		End If
		
		set rs = server.CreateObject("adodb.recordset")
		rs.open "Select * from keyicms_News Where "&Flag&datawhere, conn, 1, 2
		rs.pagesize = lNum
		totalpage = rs.pagecount
		rs.close : set rs = nothing
		If totalpage = 0 then totalpage = totalpage + 1
		
		Call html_To_NewsSort(ID, "Index."&HTMLName, "", Ke01)
		For n = 1 To totalpage
			Call html_To_NewsSort(ID, "Index"&Separated&n&"."&HTMLName, n, Ke01)
		Next

	Case "Product"
		Call html_To_Product(ID, Ke01)
	Case "ProductSort"
		set strDictionary = server.CreateObject("SCRIPTING.DICTIONARY")
		Set regExpObj = New Regexp
		regExpObj.IgnoreCase = True 
		regExpObj.Global = True
		regExpObj.Pattern="{keyicms:ProductList([\s\S]*?)}([\s\S]*?){/keyicms:ProductList}"
		replace_code = ReadFromUTF(TemplatePath&"/html/ProductList.html","utf-8")
		set matches = regExpObj.Execute(replace_code)
		for each match in matches
			labelStr = match.SubMatches(0)
			set labelArr = parseArr(labelStr)
			lNum = labelArr("Num")
		next
		set matches = nothing
		Set regExpObj = nothing
		set strDictionary = nothing

		If ID <> "" Then
			SortPath = Conn.execute("Select * from keyicms_ProductSort where ID="&ID)("SortPath")
			datawhere =  " and instr(SortPath,'"&SortPath&"')>0"
		End If
		
		set rs = server.CreateObject("adodb.recordset")
		rs.open "Select * from keyicms_Product Where "&Flag&datawhere, conn, 1, 2
		rs.pagesize = lNum
		totalpage = rs.pagecount
		rs.close : set rs = nothing
		If totalpage = 0 then totalpage = totalpage + 1
		
		Call html_To_ProductSort(ID, "Index."&HTMLName, "", Ke01)
		For n = 1 To totalpage
			Call html_To_ProductSort(ID, "Index"&Separated&n&"."&HTMLName, n, Ke01)
		Next

	Case "Index"
		Call html_To_Index(FileName, Ke01)
End Select
%>