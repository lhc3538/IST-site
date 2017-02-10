<%
'生成企业信息静态方法
Function html_To_About(ID, Ke01)
	Call SiteInfo()
	Set rs = server.CreateObject("adodb.recordset")
	sql = "select * from keyicms_About where ID="&ID
	rs.open sql, conn, 1, 2
	AboutName = rs("AboutName")
	ClassSeo = rs("ClassSeo")
	Content = rs("Content")
	AddTime = rs("AddTime")
	UpdateTime = rs("UpdateTime")
	SeoTitle = rs("AboutName")
	If rs("SeoKeywords") <> "" Then
		SeoKeywords = rs("SeoKeywords")
	Else
		SeoKeywords = rs("AboutName")
	End If
	If rs("SeoDescription") <> "" Then
		SeoDescription = rs("SeoDescription")
	Else
		SeoDescription = rs("AboutName")
	End If
	rs.close : set rs = nothing
	If Ke01 = "PC" Then Flag = "ViewFlag"
	If Ke01 = "Mobile" Then Flag = "MobileFlag"

	set rs = server.CreateObject("adodb.recordset")
	rs.open "select top 1 * from keyicms_About where ID<" & ID & " and "&Flag&" Order by ID desc", Conn, 1, 2
	if rs.eof then
		PrevLink = "#" : PrevName = "暂无"
	else
		PrevLink = rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName : PrevName = rs("AboutName")
	end if
	rs.close
	rs.open "select top 1 * from keyicms_About where ID>" & ID & " and "&Flag&" Order by ID asc", Conn, 1, 2
	if rs.eof then
		NextLink = "#" : NextName = "暂无"
	else
		NextLink = rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName : NextName = rs("AboutName")
	end if
	rs.close : set rs = nothing

	If Ke01 = "PC" Then
		replace_code = ReadFromUTF(TemplatePath&"/html/AboutView.html","utf-8")
		FilePath = "/"&AboutDir&"/"&ClassSeo&Separated&ID&"."&HTMLName&""
		TemplatePath = TemplatePath
	End If
	If Ke01 = "Mobile" Then
		replace_code = ReadFromUTF(mTemplatePath&"/html/AboutView.html","utf-8")
		FilePath = "/m/"&AboutDir&"/"&ClassSeo&Separated&ID&"."&HTMLName&""
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