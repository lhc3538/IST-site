<%
'生成产品详细静态方法
Function html_To_Product(ID, Ke01)
	Call SiteInfo()
	
	Set rs = server.CreateObject("adodb.recordset")
	rs.open "select * from keyicms_Product where ID="&ID,conn,1,2
	ProductName = rs("ProductName")
	ClassSeo = rs("ClassSeo")
	AddTime = rs("AddTime")
	UpdateTime = rs("UpdateTime")
	ProductNo = rs("ProductNo")
	Maker = rs("Maker")
	SmallPic = rs("SmallPic")
	BigPic = rs("BigPic")
	OtherPic = rs("OtherPic")
	Content = rs("Content")
	Attribute = rs("Attribute")
	Attribute_value = rs("Attribute_value")
	
	SortID = rs("SortID")
	SortName = Conn.execute("select * from keyicms_ProductSort where ID="&rs("SortID"))("SortName")
	
	If rs("SeoTitle") <> "" Then
		SeoTitle = rs("SeoTitle")
	Else
		SeoTitle = rs("ProductName")
	End If
	If rs("SeoKeywords") <> "" Then
		SeoKeywords = rs("SeoKeywords")
	Else
		SeoKeywords = rs("ProductName")
	End If
	If rs("SeoDescription") <> "" Then
		SeoDescription = rs("SeoDescription")
	Else
		SeoDescription = rs("ProductName")
	End If
	rs.close : set rs = nothing
	If Ke01 = "PC" Then Flag = "ViewFlag"
	If Ke01 = "Mobile" Then Flag = "MobileFlag"

	set rs = server.CreateObject("adodb.recordset")
	rs.open "select top 1 * from keyicms_Product where ID<" & ID & " and "&Flag&" Order by ID desc", Conn, 1, 2
	if rs.eof then
		PrevLink = "#" : PrevName = "暂无"
	else
		PrevLink = rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName : PrevName = rs("ProductName")
	end if
	rs.close
	rs.open "select top 1 * from keyicms_Product where ID>" & ID & " and "&Flag&" Order by ID asc", Conn, 1, 2
	if rs.eof then
		NextLink = "#" : NextName = "暂无"
	else
		NextLink = rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName : NextName = rs("ProductName")
	end if
	rs.close : set rs = nothing

	If Ke01 = "PC" Then
		replace_code = ReadFromUTF(TemplatePath&"/html/ProductView.html","utf-8")
		FilePath = "/"&ProDir&"/"&ClassSeo&Separated&ID&"."&HTMLName&""
		TemplatePath = TemplatePath
	End If
	If Ke01 = "Mobile" Then
		replace_code = ReadFromUTF(mTemplatePath&"/html/ProductView.html","utf-8")
		FilePath = "/m/"&ProDir&"/"&ClassSeo&Separated&ID&"."&HTMLName&""
		TemplatePath = mTemplatePath
	End If

	If SiteStatus Then
		replace_code = Position(replace_code, SortID, 3, Ke01)
	
		replace_code = ReplaceCommon(replace_code, TemplatePath, Ke01)
		
		replace_code = replace(replace_code, "{keyicms:ID}", ID)
		replace_code = replace(replace_code, "{keyicms:ProductName}", ProductName)
		replace_code = replace(replace_code, "{keyicms:ProductNo}", ProductNo)
		replace_code = replace(replace_code, "{keyicms:Maker}", Maker)
		replace_code = replace(replace_code, "{keyicms:BigPic}", BigPic)
		replace_code = replace(replace_code, "{keyicms:SmallPic}", SmallPic)
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

		'======更多属性
		If IsNull(Attribute) Then Attribute = ""
		If IsNull(Attribute_value) Then Attribute_value = ""
		AttributeShop = split(Attribute, "§§§")
		Attribute_valueShop = split(Attribute_value, "§§§")
		Set regExpObj = New Regexp
		regExpObj.IgnoreCase = True
		regExpObj.Global = True
		regExpObj.Pattern = "{keyicms:Attribute([\s\S]*?)}([\s\S]*?){/keyicms:Attribute}"
		set matches = regExpObj.Execute(replace_code)
		
		for each match in matches
			loopstrTotal = ""
			labelStr = match.SubMatches(0)
			loopStr = match.SubMatches(1)
			
			set regExpTwo = New Regexp
			regExpTwo.IgnoreCase = True
			regExpTwo.Global = True
			regExpTwo.Pattern = "{Attribute:([\s\S]*?)}"
			set matchesfield = regExpTwo.Execute(loopStr)
			loopstrTotal = ""
			For i = 0 to Ubound(AttributeShop)
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
						Case "Attribute"
							nloopstr = replace(nloopstr, matchfield.value, AttributeShop(i))
						Case "Attribute_value"
							nloopstr = replace(nloopstr, matchfield.value, Attribute_valueShop(i))
					end select
				next
				loopstrTotal = loopstrTotal & nloopstr
			Next
			replace_code = replace(replace_code, match.value, loopstrTotal)
		next
		
		'======更多图片
		If IsNull(OtherPic) Then OtherPic = ""
		OtherShop = split(OtherPic, "|")
		Set regExpObj = New Regexp
		regExpObj.IgnoreCase = True
		regExpObj.Global = True
		regExpObj.Pattern = "{keyicms:OtherPic([\s\S]*?)}([\s\S]*?){/keyicms:OtherPic}"
		set matches = regExpObj.Execute(replace_code)
		
		for each match in matches
			labelStr = match.SubMatches(0)
			loopStr = match.SubMatches(1)
			
			set regExpTwo = New Regexp
			regExpTwo.IgnoreCase = True
			regExpTwo.Global = True
			regExpTwo.Pattern = "{OtherPic:([\s\S]*?)}"
			set matchesfield = regExpTwo.Execute(loopStr)
			loopstrTotal = ""
			For i = 0 to Ubound(OtherShop)
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
							nloopstr = replace(nloopstr, matchfield.value, OtherShop(i))
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