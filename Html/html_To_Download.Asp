<%
'生成下载详细静态方法
Function html_To_Download(ID, Ke01)
	Call SiteInfo()
	
	Set rs = server.CreateObject("adodb.recordset")
	rs.open "select * from keyicms_Download where ID="&ID,conn,1,2
	DownName = rs("DownName")
	ClassSeo = rs("ClassSeo")
	Content = rs("Content")
	AddTime = rs("AddTime")
	UpdateTime = rs("UpdateTime")
	FileSize = rs("FileSize")
	FileUrl = rs("FileUrl")
	
	SortID = rs("SortID")
	SortName = Conn.execute("select * from keyicms_DownSort where ID="&rs("SortID"))("SortName")
	
	SeoTitle = rs("DownName")
	If rs("SeoKeywords") <> "" Then
		SeoKeywords = rs("SeoKeywords")
	Else
		SeoKeywords = rs("DownName")
	End If
	If rs("SeoDescription") <> "" Then
		SeoDescription = rs("SeoDescription")
	Else
		SeoDescription = rs("DownName")
	End If
	rs.close : set rs = nothing
	If Ke01 = "PC" Then Flag = "ViewFlag"
	If Ke01 = "Mobile" Then Flag = "MobileFlag"

	set rs = server.CreateObject("adodb.recordset")
	rs.open "select top 1 * from keyicms_Download where ID<" & ID & " and "&Flag&" Order by ID desc", Conn, 1, 2
	if rs.eof then
		PrevLink = "#" : PrevName = "没有了"
	else
		PrevLink = rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName : PrevName = rs("DownName")
	end if
	rs.close
	rs.open "select top 1 * from keyicms_Download where ID>" & ID & " and "&Flag&" Order by ID asc", Conn, 1, 2
	if rs.eof then
		NextLink = "#" : NextName = "没有了"
	else
		NextLink = rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName : NextName = rs("DownName")
	end if
	rs.close : set rs = nothing

	If Ke01 = "PC" Then
		replace_code = ReadFromUTF(TemplatePath&"/html/DownView.html","utf-8")
		FilePath = "/"&DownDir&"/"&ClassSeo&Separated&ID&"."&HTMLName&""
		TemplatePath = TemplatePath
	End If
	If Ke01 = "Mobile" Then
		replace_code = ReadFromUTF(mTemplatePath&"/html/DownView.html","utf-8")
		FilePath = "/m/"&DownDir&"/"&ClassSeo&Separated&ID&"."&HTMLName&""
		TemplatePath = mTemplatePath
	End If

	If SiteStatus Then
		replace_code = Position(replace_code, SortID, 5, Ke01)
	
		replace_code = ReplaceCommon(replace_code, TemplatePath, Ke01)
		
		replace_code = replace(replace_code, "{keyicms:ID}", ID)
		replace_code = replace(replace_code, "{keyicms:DownName}", DownName)
		replace_code = replace(replace_code, "{keyicms:FileSize}", FileSize)
		replace_code = replace(replace_code, "{keyicms:FileUrl}", FileUrl)
		replace_code = replace(replace_code, "{keyicms:SortID}", SortID)
		replace_code = replace(replace_code, "{keyicms:SortName}", SortName)
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