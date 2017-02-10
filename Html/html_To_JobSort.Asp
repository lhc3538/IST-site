<%
'生成招聘列表静态方法
Function html_To_JobSort(FileName, Page, ke01)
	Call SiteInfo()
	set rs = server.CreateObject("adodb.recordset")
	rs.open "select * from keyicms_System where DataForm=6", conn, 1, 2
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
		If Checkfile(TemplatePath&"/html/Job.html") Then
			replace_code = ReadFromUTF(TemplatePath&"/html/Job.html","utf-8")
		Else
			replace_code = ReadFromUTF(TemplatePath&"/html/JobList.html","utf-8")
		End If
		FilePath = "/" & JobDir & FolderName & "/" & FileName
		TemplatePath = TemplatePath
	End If
	If Ke01 = "Mobile" Then
		If Checkfile(mTemplatePath&"/html/Job.html") Then
			replace_code = ReadFromUTF(mTemplatePath&"/html/Job.html","utf-8")
		Else
			replace_code = ReadFromUTF(mTemplatePath&"/html/JobList.html","utf-8")
		End If
		FilePath = "/m/" & JobDir & FolderName & "/" & FileName
		TemplatePath = mTemplatePath
	End If

	If SiteStatus Then
		replace_code = Position(replace_code, "", 6, Ke01)
	
		replace_code = ReplaceCommon(replace_code, TemplatePath, Ke01)

		replace_code = replace(replace_code, "{keyicms:SortName}", SortName)
		replace_code = replace(replace_code, "{keyicms:SeoTitle}", SeoTitle)
		replace_code = replace(replace_code, "{keyicms:SeoKeywords}", SeoKeywords)
		replace_code = replace(replace_code, "{keyicms:SeoDescription}", SeoDescription)

		replace_code = HtmlTo(replace_code, "", "", Ke01, "")
	Else
		replace_code = ExplainJS
	End If
	
	If Ke01 = "PC" Then replace_code = replace_code & ConstJs
	Call WriteToUTF(FilePath, replace_code)
End Function
%>