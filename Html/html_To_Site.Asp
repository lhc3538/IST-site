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
'生成全站静态方法
Call SiteInfo()
Table = request("Table")
Ke01 = request("Ke01")
If Ke01 = "PC" Then
	Flag = " ViewFlag "
	TemplatePath = TemplatePath
End If
If Ke01 = "Mobile" Then
	Flag = " MobileFlag "
	TemplatePath = mTemplatePath
End If

set rsSite = server.CreateObject("adodb.recordset")
rsSite.open "select * from keyicms_System where DataForm=1", conn, 1, 2
If rsSite("ViewFlag") Then
	ErrMsg = ErrMsg & "======"&rsSite("SystemName")&"=======<br />"
	If Checkfile(TemplatePath&"/html/AboutView.html") Then
		Set rshtml = Conn.Execute("Select * from keyicms_About where "&Flag)
		do while not rshtml.eof
			Call html_To_About(rshtml("ID"), Ke01)
			rshtml.movenext
		loop
		rshtml.close : set rshtml = Nothing
		ErrMsg = ErrMsg & rsSite("SystemName") & "的详细页生成成功<br />"
	Else
		ErrMsg = ErrMsg & "<font style='color:#F00'>" & rsSite("SystemName") & "的AboutView.html模板未找到，详细页未生成</font><br />"
	End If

	If not Checkfile(TemplatePath&"/html/About.html") and not Checkfile(TemplatePath&"/html/AboutView.html") Then
		ErrMsg = ErrMsg & "<font style='color:#F00'>" & rsSite("SystemName") & "的About.html和AboutView.html模板均未找到，栏目页未生成</font><br />"
	Else
		Call html_To_AboutSort("Index."&HTMLName, Ke01)
		ErrMsg = ErrMsg & rsSite("SystemName") & "的栏目页生成成功<br />"
	End If
End If
rsSite.close : set rsSite = nothing

set rsSite = server.CreateObject("adodb.recordset")
rsSite.open "select * from keyicms_System where DataForm=2", conn, 1, 2
If rsSite("ViewFlag") Then
	ErrMsg = ErrMsg & "======"&rsSite("SystemName")&"=======<br />"
	If Checkfile(TemplatePath&"/html/NewsView.html") Then
		Set rshtml = Conn.Execute("Select * from keyicms_News where "&Flag)
		do while not rshtml.eof
			Call html_To_News(rshtml("ID"), Ke01)	
			rshtml.movenext
		loop
		rshtml.close : set rshtml = Nothing		
		ErrMsg = ErrMsg & rsSite("SystemName") & "的详细页生成成功<br />"
	Else
		ErrMsg = ErrMsg & "<font style='color:#F00'>" & rsSite("SystemName") & "的NewsView.html模板未找到，详细页未生成</font><br />"
	End If
	
	If Checkfile(TemplatePath&"/html/NewsList.html") Then
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
		
		set rs = server.CreateObject("adodb.recordset")
		rs.open "Select * from keyicms_News Where "&Flag, conn, 1, 2
		rs.pagesize = lNum
		totalpage = rs.pagecount
		rs.close : set rs = nothing
		If totalpage = 0 then totalpage = totalpage + 1
		
		For n = 1 To totalpage
			Call html_To_NewsSort("", "Index"&Separated&n&"."&HTMLName, n, Ke01)
		Next
		
		Set rshtml = Server.CreateObject("ADODB.Recordset")
		sql = "select * from keyicms_NewsSort order by ID desc"
		rshtml.Open sql, conn, 1, 1
		If Not rshtml.EOF Then
			Do While Not rshtml.EOF
				ID = rshtml("ID")
				set rs = server.CreateObject("adodb.recordset")
				rs.open "Select * from keyicms_News where "&Flag&" and instr(SortPath,'"&rshtml("SortPath")&"')>0",conn,1,2
				rs.pagesize = lNum
				totalpage = rs.pagecount
				rs.close : set rs = nothing
				If totalpage = 0 then totalpage = totalpage + 1
				Call html_To_NewsSort(ID, "Index."&HTMLName, 1, Ke01)
				For n = 1 To totalpage
					Call html_To_NewsSort(ID, "Index"&Separated&n&"."&HTMLName, n, Ke01)
				Next
				rshtml.movenext
			Loop
		End If
		rshtml.Close : Set rshtml = Nothing
		ErrMsg = ErrMsg & rsSite("SystemName") & "的列表页生成成功<br />"
	Else
		ErrMsg = ErrMsg & "<font style='color:#F00'>" & rsSite("SystemName") & "的NewsList.html模板未找到，列表页未生成</font><br />"
	End If

	If not Checkfile(TemplatePath&"/html/News.html") and not Checkfile(TemplatePath&"/html/NewsList.html") Then
		ErrMsg = ErrMsg & "<font style='color:#F00'>" & rsSite("SystemName") & "News.html和NewsList.html模板均未找到，栏目页未生成</font><br />"
	Else
		Call html_To_NewsSort("", "Index."&HTMLName, "", Ke01)
		ErrMsg = ErrMsg & rsSite("SystemName") & "的栏目页生成成功<br />"
	End If
End If
rsSite.close : set rsSite = nothing

'====================生成Product
set rsSite = server.CreateObject("adodb.recordset")
rsSite.open "select * from keyicms_System where DataForm=3", conn, 1, 2
If rsSite("ViewFlag") Then
	ErrMsg = ErrMsg & "======"&rsSite("SystemName")&"=======<br />"
	If Checkfile(TemplatePath&"/html/ProductView.html") Then
		Set rshtml = Conn.Execute("Select * from keyicms_Product where "&Flag)
		do while not rshtml.eof
			Call html_To_Product(rshtml("ID"), Ke01)	
			rshtml.movenext
		loop
		rshtml.close : set rshtml = Nothing		
		ErrMsg = ErrMsg & rsSite("SystemName") & "的详细页生成成功<br />"
	Else
		ErrMsg = ErrMsg & "<font style='color:#F00'>" & rsSite("SystemName") & "的ProductView.html模板未找到，详细页未生成</font><br />"
	End If

	If Checkfile(TemplatePath&"/html/ProductList.html") Then
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
		
		set rs = server.CreateObject("adodb.recordset")
		rs.open "Select * from keyicms_Product Where "&Flag, conn, 1, 2
		rs.pagesize = lNum
		totalpage = rs.pagecount
		rs.close : set rs = nothing
		If totalpage = 0 then totalpage = totalpage + 1
		
		For n = 1 To totalpage
			Call html_To_ProductSort("", "Index"&Separated&n&"."&HTMLName, n, Ke01)
		Next
		
		Set rshtml = Server.CreateObject("ADODB.Recordset")
		sql = "select * from keyicms_ProductSort order by ID desc"
		rshtml.Open sql, conn, 1, 1
		
		If Not rshtml.EOF Then
			Do While Not rshtml.EOF
				ID = rshtml("ID")
				set rs = server.CreateObject("adodb.recordset")
				rs.open "Select * from keyicms_Product where "&Flag&" and instr(SortPath,'"&rshtml("SortPath")&"')>0",conn,1,2
				rs.pagesize = lNum
				totalpage = rs.pagecount
				rs.close : set rs = nothing
				If totalpage = 0 then totalpage = totalpage + 1
				Call html_To_ProductSort(ID, "Index."&HTMLName, 1, Ke01)
				For n = 1 To totalpage
					Call html_To_ProductSort(ID, "Index"&Separated&n&"."&HTMLName, n, Ke01)
				Next
				rshtml.movenext
			Loop
		End If
		rshtml.Close : Set rshtml = Nothing
		ErrMsg = ErrMsg & rsSite("SystemName") & "的列表页生成成功<br />"
	Else
		ErrMsg = ErrMsg & "<font style='color:#F00'>" & rsSite("SystemName") & "的ProductList.html模板未找到，列表页未生成</font><br />"
	End If

	If not Checkfile(TemplatePath&"/html/Product.html") and not Checkfile(TemplatePath&"/html/ProductList.html") Then
		ErrMsg = ErrMsg & "<font style='color:#F00'>" & rsSite("SystemName") & "的Product.html和ProductList.html模板均未找到，栏目页未生成</font><br />"
	Else
		Call html_To_ProductSort("", "Index."&HTMLName, "", Ke01)
		ErrMsg = ErrMsg & rsSite("SystemName") & "的栏目页生成成功<br />"
	End If
End If
rsSite.close : set rsSite = nothing

set rsSite = server.CreateObject("adodb.recordset")
rsSite.open "select * from keyicms_System where DataForm=4", conn, 1, 2
If rsSite("ViewFlag") Then
	ErrMsg = ErrMsg & "======"&rsSite("SystemName")&"=======<br />"
	If Checkfile(TemplatePath&"/html/CaseView.html") Then
		Set rshtml = Conn.Execute("Select * from keyicms_Case where "&Flag)
		do while not rshtml.eof
			Call html_To_Case(rshtml("ID"), Ke01)	
			rshtml.movenext
		loop
		rshtml.close : set rshtml = Nothing
		ErrMsg = ErrMsg & rsSite("SystemName") & "的详细页生成成功<br />"
	Else
		ErrMsg = ErrMsg & "<font style='color:#F00'>" & rsSite("SystemName") & "的CaseView.html模板未找到，详细页未生成</font><br />"
	End If

	If Checkfile(TemplatePath&"/html/CaseList.html") Then
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
		
		set rs = server.CreateObject("adodb.recordset")
		rs.open "Select * from keyicms_Case Where "&Flag, conn, 1, 2
		rs.pagesize = lNum
		totalpage = rs.pagecount
		rs.close : set rs = nothing
		If totalpage = 0 then totalpage = totalpage + 1
		
		For n = 1 To totalpage
			Call html_To_CaseSort("", "Index"&Separated&n&"."&HTMLName, n, Ke01)
		Next
	
		Set rshtml = Server.CreateObject("ADODB.Recordset")
		sql = "select * from keyicms_CaseSort order by ID desc"
		rshtml.Open sql, conn, 1, 1
		If Not rshtml.EOF Then
			Do While Not rshtml.EOF
				ID = rshtml("ID")
				set rs = server.CreateObject("adodb.recordset")
				rs.open "Select * from keyicms_Case where "&Flag&" and instr(SortPath,'"&rshtml("SortPath")&"')>0",conn,1,2
				rs.pagesize = lNum
				totalpage = rs.pagecount
				rs.close : set rs = nothing
				If totalpage = 0 then totalpage = totalpage + 1
				Call html_To_CaseSort(ID, "Index."&HTMLName, 1, Ke01)
				For n = 1 To totalpage
					Call html_To_CaseSort(ID, "Index"&Separated&n&"."&HTMLName, n, Ke01)
				Next
				rshtml.movenext
			Loop
		End If
		rshtml.Close : Set rshtml = Nothing
		ErrMsg = ErrMsg & rsSite("SystemName") & "的列表页生成成功<br />"
	Else
		ErrMsg = ErrMsg & "<font style='color:#F00'>" & rsSite("SystemName") & "的CaseList.html模板未找到，列表页未生成</font><br />"
	End If

	If not Checkfile(TemplatePath&"/html/Case.html") and not Checkfile(TemplatePath&"/html/CaseList.html") Then
		ErrMsg = ErrMsg & "<font style='color:#F00'>" & rsSite("SystemName") & "的Case.html和CaseList.html模板未找到，栏目页未成功</font><br />"
	Else
		Call html_To_CaseSort("", "Index."&HTMLName, "", Ke01)
		ErrMsg = ErrMsg & rsSite("SystemName") & "的栏目页生成成功<br />"
	End If
End If
rsSite.close : set rsSite = nothing

set rsSite = server.CreateObject("adodb.recordset")
rsSite.open "select * from keyicms_System where DataForm=5", conn, 1, 2
If rsSite("ViewFlag") Then
	ErrMsg = ErrMsg & "======"&rsSite("SystemName")&"=======<br />"
	If Checkfile(TemplatePath&"/html/DownView.html") Then
		Set rshtml = Conn.Execute("Select * from keyicms_Download where "&Flag)
		do while not rshtml.eof
			Call html_To_Download(rshtml("ID"), Ke01)	
			rshtml.movenext
		loop
		rshtml.close : set rshtml = Nothing		
		ErrMsg = ErrMsg & rsSite("SystemName") & "的详细页生成成功<br />"
	Else
		ErrMsg = ErrMsg & "<font style='color:#F00'>" & rsSite("SystemName") & "的DownView.html模板未找到，详细页未生成</font><br />"
	End If

	If Checkfile(TemplatePath&"/html/DownList.html") Then
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
		
		set rs = server.CreateObject("adodb.recordset")
		rs.open "Select * from keyicms_Download Where "&Flag, conn, 1, 2
		rs.pagesize = lNum
		totalpage = rs.pagecount
		rs.close : set rs = nothing
		If totalpage = 0 then totalpage = totalpage + 1
		
		For n = 1 To totalpage
			Call html_To_DownSort("", "Index"&Separated&n&"."&HTMLName, n, Ke01)
		Next
		
		Set rshtml = Server.CreateObject("ADODB.Recordset")
		sql = "Select * from keyicms_DownSort Order by ID Desc"
		rshtml.Open sql, conn, 1, 1
		If Not rshtml.EOF Then
			Do While Not rshtml.EOF
				ID = rshtml("ID")
				set rs = server.CreateObject("adodb.recordset")
				rs.open "Select * from keyicms_Download Where "&Flag&" and instr(SortPath,'"&rshtml("SortPath")&"')>0",conn,1,2
				rs.pagesize = lNum
				totalpage = rs.pagecount
				rs.close : set rs = nothing
				If totalpage = 0 then totalpage = totalpage + 1
				Call html_To_DownSort(ID, "Index."&HTMLName, 1, Ke01)
				For n = 1 To totalpage
					Call html_To_DownSort(ID, "Index"&Separated&n&"."&HTMLName, n, Ke01)
				Next
				rshtml.movenext
			Loop
		End If
		rshtml.Close : Set rshtml = Nothing
		ErrMsg = ErrMsg & rsSite("SystemName") & "的列表页生成成功<br />"
	Else
		ErrMsg = ErrMsg & "<font style='color:#F00'>" & rsSite("SystemName") & "的DownList.html模板未找到，列表页未生成</font><br />"
	End If

	If not Checkfile(TemplatePath&"/html/Download.html") and not Checkfile(TemplatePath&"/html/DownList.html") Then
		ErrMsg = ErrMsg & "<font style='color:#F00'>" & rsSite("SystemName") & "的Download.html和DownList.html模板未找到，未生成</font><br />"
	Else
		ErrMsg = ErrMsg & rsSite("SystemName") & "的栏目页生成成功<br />"
		Call html_To_DownSort("", "Index."&HTMLName, "", Ke01)
	End If
End If
rsSite.close : set rsSite = nothing

set rsSite = server.CreateObject("adodb.recordset")
rsSite.open "select * from keyicms_System where DataForm=6", conn, 1, 2
If rsSite("ViewFlag") Then
	ErrMsg = ErrMsg & "======"&rsSite("SystemName")&"=======<br />"
	If Checkfile(TemplatePath&"/html/JobView.html") Then
		Set rshtml = Conn.Execute("Select * from keyicms_Job where "&Flag)
		do while not rshtml.eof
			Call html_To_Job(rshtml("ID"), Ke01)	
			rshtml.movenext
		loop
		rshtml.close : set rshtml = Nothing		
		ErrMsg = ErrMsg & rsSite("SystemName") & "的详细页生成成功<br />"
	Else
		ErrMsg = ErrMsg & "<font style='color:#F00'>" & rsSite("SystemName") & "的JobView.html模板未找到，详细未生成</font><br />"
	End If
	
	If Checkfile(TemplatePath&"/html/JobList.html") Then
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
		rs.open "Select * from keyicms_Job Where "&Flag, conn, 1, 2
		rs.pagesize = lNum
		totalpage = rs.pagecount
		rs.close : set rs = nothing
		If totalpage = 0 then totalpage = totalpage + 1
		
		For n = 1 To totalpage
			Call html_To_JobSort("Index"&Separated&n&"."&HTMLName, n, Ke01)
		Next
		ErrMsg = ErrMsg & rsSite("SystemName") & "的列表页生成成功<br />"
	Else
		ErrMsg = ErrMsg & "<font style='color:#F00'>" & rsSite("SystemName") & "的JobList.html模板未找到，列表页未生成</font><br />"
	End If

	If not Checkfile(TemplatePath&"/html/Job.html") and not Checkfile(TemplatePath&"/html/JobList.html") Then
		ErrMsg = ErrMsg & "<font style='color:#F00'>" & rsSite("SystemName") & "的Job.html和JobList.html模板未找到，未生成</font><br />"
	Else
		Call html_To_JobSort("Index."&HTMLName, "", Ke01)
		ErrMsg = ErrMsg & rsSite("SystemName") & "的栏目页生成成功<br />"
	End If
End If
rsSite.close : set rsSite = nothing

If Checkfile(TemplatePath&"/html/Index.html") Then
	Call html_To_Index("Index", Ke01)
	ErrMsg = ErrMsg & "<br />首页生成成功<br />"
Else
	ErrMsg = ErrMsg & "<font style='color:#F00'>首页模板未找到，未生成</font><br />"
End If
If Checkfile(TemplatePath&"/html/Contact.html") Then
	Call html_To_Index("Contact", Ke01)
	ErrMsg = ErrMsg & "联系生成成功<br />"
Else
	ErrMsg = ErrMsg & "<font style='color:#F00'>首页的模板未找到，未生成</font><br />"
End If
If Checkfile(TemplatePath&"/html/Message.html") Then
	Call html_To_Index("Message", Ke01)
	ErrMsg = ErrMsg & "留言生成成功<br />"
Else
	ErrMsg = ErrMsg & "<font style='color:#F00'>留言的模板未找到，未生成</font><br />"
End If

response.Write ErrMsg
%>