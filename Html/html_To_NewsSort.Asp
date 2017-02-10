<%
'生成新闻列表静态方法
Function html_To_NewsSort(ID, FileName, Page, ke01)
	Call SiteInfo()
	If ID = "" Then
		set rs = server.CreateObject("adodb.recordset")
		rs.open "select * from keyicms_System where DataForm=2", conn, 1, 2
		SortName = rs("SeoTitle")
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
		Content = ""
		SortPath = "0,"
		FolderName = ""
		rs.close : set rs = nothing

		If Ke01 = "PC" Then
			If Checkfile(TemplatePath&"/html/News.html") and Page = "" Then
				replace_code = ReadFromUTF(TemplatePath&"/html/News.html","utf-8")
			Else
				replace_code = ReadFromUTF(TemplatePath&"/html/NewsList.html","utf-8")
			End If
		End If
		If Ke01 = "Mobile" Then
			If Checkfile(mTemplatePath&"/html/News.html") and Page = "" Then
				replace_code = ReadFromUTF(mTemplatePath&"/html/News.html","utf-8")
			Else
				replace_code = ReadFromUTF(mTemplatePath&"/html/NewsList.html","utf-8")
			End If
		End If
	Else
		Set rs = server.CreateObject("adodb.recordset")
		rs.open "select * from keyicms_NewsSort where ID="&ID, conn, 1, 2
		SortName = rs("SortName")
		SortPath = rs("SortPath")
		Content = changeNull(rs("Content"))

		SeoTitle = rs("SortName")
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
		FolderName = "/" & rs("FolderName")
		rs.close : set rs = nothing

		If Ke01 = "PC" Then replace_code = ReadFromUTF(TemplatePath&"/html/NewsList.html","utf-8")
		If Ke01 = "Mobile" Then replace_code = ReadFromUTF(mTemplatePath&"/html/NewsList.html","utf-8")
	End If

	If Ke01 = "PC" Then
		FilePath = "/" & NewsDir & FolderName & "/" & FileName
		TemplatePath = TemplatePath
	End If
	If Ke01 = "Mobile" Then
		FilePath = "/m/" & NewsDir & FolderName & "/" & FileName
		TemplatePath = mTemplatePath
	End If

	If SiteStatus Then
		replace_code = Position(replace_code, ID, 2, Ke01)
	
		replace_code = replace(replace_code, "{keyicms:SortName}", SortName)
		replace_code = ReplaceCommon(replace_code, TemplatePath, Ke01)

		replace_code = replace(replace_code, "{keyicms:ID}", ID)
		replace_code = replace(replace_code, "{keyicms:SeoTitle}", SeoTitle)
		replace_code = replace(replace_code, "{keyicms:SeoKeywords}", SeoKeywords)
		replace_code = replace(replace_code, "{keyicms:SeoDescription}", SeoDescription)
		replace_code = replace(replace_code, "{keyicms:Content}", Content)
		replace_code = HtmlTo(replace_code, "News", Page, Ke01, SortPath)
	Else
		replace_code = ExplainJS
	End If
	
	If Ke01 = "PC" Then replace_code = replace_code & ConstJs
	Call WriteToUTF(FilePath, replace_code)
End Function
%>