<%
'生成案例列表静态方法
Function html_To_Case(ID, ke01)
	Call SiteInfo()
	
	Set rs = server.CreateObject("adodb.recordset")
	rs.open "select * from keyicms_Case where ID="&ID,conn,1,2
	CaseName = rs("CaseName")
	ClassSeo = rs("ClassSeo")
	Content = rs("Content")
	AddTime = rs("AddTime")
	UpdateTime = rs("UpdateTime")
	SmallPic = rs("SmallPic")
	BigPic = rs("BigPic")
	OtherPic = rs("OtherPic")
	
	SortID = rs("SortID")
	SortName = Conn.execute("select * from keyicms_CaseSort where ID="&rs("SortID"))("SortName")
	
	If rs("SeoTitle") <> "" Then
		SeoTitle = rs("SeoTitle")
	Else
		SeoTitle = rs("CaseName")
	End If
	If rs("SeoKeywords") <> "" Then
		SeoKeywords = rs("SeoKeywords")
	Else
		SeoKeywords = rs("CaseName")
	End If
	If rs("SeoDescription") <> "" Then
		SeoDescription = rs("SeoDescription")
	Else
		SeoDescription = rs("CaseName")
	End If
	rs.close : set rs = nothing
	If Ke01 = "PC" Then Flag = "ViewFlag"
	If Ke01 = "Mobile" Then Flag = "MobileFlag"

	set rs = server.CreateObject("adodb.recordset")
	rs.open "select top 1 * from keyicms_Case where ID<" & ID & " and "&Flag&" Order by ID desc", Conn, 1, 2
	if rs.eof then
		PrevLink = "#" : PrevName = "暂无"
	else
		PrevLink = rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName : PrevName = rs("CaseName")
	end if
	rs.close
	rs.open "select top 1 * from keyicms_Case where ID>" & ID & " and "&Flag&" Order by ID asc", Conn, 1, 2
	if rs.eof then
		NextLink = "#" : NextName = "暂无"
	else
		NextLink = rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName : NextName = rs("CaseName")
	end if
	rs.close : set rs = nothing

	If Ke01 = "PC" Then
		replace_code = ReadFromUTF(TemplatePath&"/html/CaseView.html","utf-8")
		FilePath = "/"&CaseDir&"/"&ClassSeo&Separated&ID&"."&HTMLName&""
		TemplatePath = TemplatePath
	End If
	If Ke01 = "Mobile" Then
		replace_code = ReadFromUTF(mTemplatePath&"/html/CaseView.html","utf-8")
		FilePath = "/m/"&CaseDir&"/"&ClassSeo&Separated&ID&"."&HTMLName&""
		TemplatePath = mTemplatePath
	End If

	If SiteStatus Then
		replace_code = Position(replace_code, SortID, 4, Ke01)
	
		replace_code = ReplaceCommon(replace_code, TemplatePath, Ke01)
		
		replace_code = replace(replace_code, "{keyicms:ID}", ID)
		replace_code = replace(replace_code, "{keyicms:CaseName}", CaseName)
		replace_code = replace(replace_code, "{keyicms:SmallPic}", SmallPic)
		replace_code = replace(replace_code, "{keyicms:BigPic}", BigPic)
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

		'======更多图片
		OtherPicShop = split(changeNull(OtherPic), "|")
		
		Set regExpObj = New Regexp
		regExpObj.IgnoreCase = True
		regExpObj.Global = True
		regExpObj.Pattern = "{keyicms:OtherPic([\s\S]*?)}([\s\S]*?){/keyicms:OtherPic}"
		Set matches = regExpObj.Execute(replace_code)
		
		for each match in matches
			loopstrTotal = ""
			labelStr = match.SubMatches(0)
			loopStr = match.SubMatches(1)
			
			set regExpTwo = New Regexp
			regExpTwo.IgnoreCase = True
			regExpTwo.Global = True
			regExpTwo.Pattern = "{OtherPic:([\s\S]*?)}"
			set matchesfield = regExpTwo.Execute(loopStr)
			For i = 0 to Ubound(OtherPicShop)
				nloopstr = loopStr
				for each matchfield in matchesfield
					fieldNameArr = regExpReplace(matchfield.SubMatches(0),"[\s]+",chr(32))
					fieldNameArr = trim(fieldNameArr)
					m = instr(fieldNameArr,chr(32))
					if  m > 0 then 
						fieldArr=""
						fieldName = left(fieldNameArr,m - 1)
						fieldArr =	right(fieldNameArr,len(fieldNameArr) - m)
					else
						fieldName = fieldNameArr
						fieldArr =	""
					end if
					select case fieldName
						Case "OtherPic"
							nloopstr = replace(nloopstr, matchfield.value, OtherPicShop(i))
					end select
				next
				loopstrTotal = loopstrTotal & nloopstr
			Next
			replace_code = replace(replace_code, match.value, loopstrTotal)
		next
		Set matches = nothing
		Set regExpObj = nothing
		
		replace_code = HtmlTo(replace_code, "", "", Ke01, "")
	Else
		replace_code = ExplainJS
	End If
	
	If Ke01 = "PC" Then replace_code = replace_code & ConstJs
	Call WriteToUTF(FilePath, replace_code)
End Function
%>