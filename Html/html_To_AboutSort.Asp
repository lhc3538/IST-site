<%
'生成企业栏目页静态方法
Function html_To_AboutSort(FileName, Ke01)
	Call SiteInfo()
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select * from keyicms_System where DataForm=1", conn, 1, 2
	SeoTitle = rs("SeoTitle")
	If rs("SeoKeywords") <> "" Then
		SeoKeywords = rs("SeoKeywords")
	Else
		SeoKeywords = SeoTitle
	End If
	If rs("SeoDescription") <> "" Then
		SeoDescription = rs("SeoDescription")
	Else
		SeoDescription = SeoTitle
	End If
	rs.close : set rs = nothing

	Set rs = server.CreateObject("adodb.recordset")
	rs.open "select * from keyicms_About where ViewFlag Order by Sequence asc", conn, 1, 2
	AboutName = rs("AboutName")
	Content = rs("Content")
	AddTime = rs("AddTime")
	UpdateTime = rs("UpdateTime")
	ID = rs("ID")
	rs.close : set rs = nothing

	set rs = server.CreateObject("adodb.recordset")
	rs.open "select top 1 * from keyicms_About where ID<" & ID & " Order by ID desc", Conn, 1, 2
	if rs.eof then
		PrevLink = "#" : PrevName = "暂无"
	else
		PrevLink = rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName : PrevName = rs("AboutName")
	end if
	rs.close
	rs.open "select top 1 * from keyicms_About where ID>" & ID & " Order by ID asc", Conn, 1, 2
	if rs.eof then
		NextLink = "#" : NextName = "暂无"
	else
		NextLink = rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName : NextName = rs("AboutName")
	end if
	rs.close : set rs = nothing

	If Ke01 = "PC" Then
		If Checkfile(TemplatePath&"/html/About.html") Then
			replace_code = ReadFromUTF(TemplatePath&"/html/About.html","utf-8")
		Else
			replace_code = ReadFromUTF(TemplatePath&"/html/AboutView.html","utf-8")
		End If
		FilePath = "/"&AboutDir & FolderName & "/" & FileName
		TemplatePath = TemplatePath
	End If
	If Ke01 = "Mobile" Then
		If Checkfile(mTemplatePath&"/html/About.html") Then
			replace_code = ReadFromUTF(mTemplatePath&"/html/About.html","utf-8")
		Else
			replace_code = ReadFromUTF(mTemplatePath&"/html/AboutView.html","utf-8")
		End If
		FilePath = "/m/"&AboutDir & FolderName & "/" & FileName
		TemplatePath = mTemplatePath
	End If

	If SiteStatus Then
		replace_code = Position(replace_code, "", 1, Ke01)
		
		replace_code = ReplaceCommon(replace_code, TemplatePath, Ke01)

		replace_code = replace(replace_code, "{keyicms:ID}", ID)
		replace_code = replace(replace_code, "{keyicms:AboutName}", AboutName)
		replace_code = replace(replace_code, "{keyicms:Content}", Content)
		replace_code = replace(replace_code, "{keyicms:SeoTitle}", SeoTitle)
		replace_code = replace(replace_code, "{keyicms:SeoKeywords}", SeoKeywords)
		replace_code = replace(replace_code, "{keyicms:SeoDescription}", SeoDescription)
		replace_code = ParseTime(replace_code, AddTime, "AddTime")
		replace_code = ParseTime(replace_code, UpdateTime, "UpdateTime")

		replace_code = replace(replace_code, "{keyicms:PrevName}", PrevName)
		replace_code = replace(replace_code, "{keyicms:PrevLink}", PrevLink)
		replace_code = replace(replace_code, "{keyicms:NextName}", NextName)
		replace_code = replace(replace_code, "{keyicms:NextLink}", NextLink)

		replace_code = HtmlTo(replace_code, "", "", Ke01, "")
	Else
		replace_code = ExplainJS
	End If

	If Ke01 = "PC" Then replace_code = replace_code & ConstJs
	Call WriteToUTF(FilePath, replace_code)
End Function
%>