<%
'生成招聘列表静态方法
Function html_To_Job(ID, ke01)
	Call SiteInfo()
	
	Set rs = server.CreateObject("adodb.recordset")
	rs.open "select * from keyicms_Job where ID="&ID,conn,1,2
	JobName = rs("JobName")
	JobAddress = rs("JobAddress")
	JobNumber = rs("JobNumber")
	Emolument = rs("Emolument")
	Requirement = rs("Requirement")
	StartDate = rs("StartDate")
	EndDate = rs("EndDate")
	ClassSeo = rs("ClassSeo")
	AddTime = rs("AddTime")
	UpdateTime = rs("UpdateTime")
	
	SeoTitle = rs("JobName")
	SeoKeywords = rs("JobName")
	SeoDescription = rs("JobName")
	rs.close : set rs = nothing
	If Ke01 = "PC" Then Flag = "ViewFlag"
	If Ke01 = "Mobile" Then Flag = "MobileFlag"

	set rs = server.CreateObject("adodb.recordset")
	rs.open "select top 1 * from keyicms_Job where ID<" & ID & " and "&Flag&" Order by ID desc", Conn, 1, 2
	if rs.eof then
		PrevLink = "#" : PrevName = "暂无"
	else
		PrevLink = rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName : PrevName = rs("JobName")
	end if
	rs.close
	rs.open "select top 1 * from keyicms_Job where ID>" & ID & " and "&Flag&" Order by ID asc", Conn, 1, 2
	if rs.eof then
		NextLink = "#" : NextName = "暂无"
	else
		NextLink = rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName : NextName = rs("JobName")
	end if
	rs.close : set rs = nothing

	If Ke01 = "PC" Then
		replace_code = ReadFromUTF(TemplatePath&"/html/JobView.html","utf-8")
		FilePath = "/"&JobDir&"/"&ClassSeo&Separated&ID&"."&HTMLName&""
		TemplatePath = TemplatePath
	End If
	If Ke01 = "Mobile" Then
		replace_code = ReadFromUTF(mTemplatePath&"/html/JobView.html","utf-8")
		FilePath = "/m/"&JobDir&"/"&ClassSeo&Separated&ID&"."&HTMLName&""
		TemplatePath = mTemplatePath
	End If

	If SiteStatus Then
		replace_code = Position(replace_code, "", 6, Ke01)
	
		replace_code = ReplaceCommon(replace_code, TemplatePath, Ke01)
		
		replace_code = replace(replace_code, "{keyicms:ID}", ID)
		replace_code = replace(replace_code, "{keyicms:JobName}", JobName)
		replace_code = replace(replace_code, "{keyicms:JobAddress}", JobAddress)
		replace_code = replace(replace_code, "{keyicms:JobNumber}", JobNumber)
		replace_code = replace(replace_code, "{keyicms:Emolument}", Emolument)
		replace_code = replace(replace_code, "{keyicms:Requirement}", Requirement)
		replace_code = ParseTime(replace_code, StartDate, "StartDate")
		replace_code = ParseTime(replace_code, EndDate, "EndDate")
		
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