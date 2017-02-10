<%
'生成静态方法公共部分
Function ReplaceCommon(replace_code, TemplatePath, Ke01)
	Dim match, matches, matchfield, matchesfield, loopstrTotal, lSortPath, lOrder, lNum, fieldName, i, labelRuleField, m, namelen
	replace_code = Include(replace_code, TemplatePath)

	replace_code = replace(replace_code, "{keyicms:TemplatePath}", TemplatePath)
	replace_code = replace(replace_code, "{keyicms:SiteTitle}", SiteTitle)
	replace_code = replace(replace_code, "{keyicms:Keywords}", Keywords)
	replace_code = replace(replace_code, "{keyicms:Descriptions}", Descriptions)
	replace_code = replace(replace_code, "{keyicms:Stats}", Stats)
	replace_code = replace(replace_code, "{keyicms:SiteUrl}", SiteUrl)
	replace_code = replace(replace_code, "{keyicms:ComName}", ComName)
	replace_code = replace(replace_code, "{keyicms:Address}", Address)
	replace_code = replace(replace_code, "{keyicms:Telephone}", Telephone)
	replace_code = replace(replace_code, "{keyicms:Mobile}", Mobile)
	replace_code = replace(replace_code, "{keyicms:Contacts}", Contacts)
	replace_code = replace(replace_code, "{keyicms:Fax}", Fax)
	replace_code = replace(replace_code, "{keyicms:Email}", Email)
	replace_code = replace(replace_code, "{keyicms:IcpNumber}", IcpNumber)
	replace_code = replace(replace_code, "{keyicms:SiteLogo}", SiteLogo)
	replace_code = replace(replace_code, "{keyicms:SiteDetail}", SiteDetail)
	replace_code = replace(replace_code, "{keyicms:Ico}", Ico)
	replace_code = replace(replace_code, "{keyicms:KeyicmsV}", KeyicmsV)

	If mStatus Then
		replace_code = replace(replace_code, "{keyicms:mTemplatePath}", mTemplatePath)
		replace_code = replace(replace_code, "{keyicms:MobileTitle}", MobileTitle)
		replace_code = replace(replace_code, "{keyicms:MobileLogo}", MobileLogo)
		replace_code = replace(replace_code, "{keyicms:MobileKeywords}", MobileKeywords)
		replace_code = replace(replace_code, "{keyicms:MobileDescriptions}", MobileDescriptions)
		replace_code = replace(replace_code, "{keyicms:MobileUrl}", MobileUrl)
		replace_code = replace(replace_code, "{keyicms:MobileTelephone}", MobileTelephone)
	End If

	replace_code = NavList(replace_code, Ke01)	
	replace_code = SortLoop(replace_code, Ke01)
	replace_code = ReplaceLoop("FriendLink", replace_code)
	replace_code = ReplaceLoop("Slide", replace_code)
	replace_code = ReplaceLoop("Ads", replace_code)
	replace_code = ReplaceLoop("Contact", replace_code)
	replace_code = Map(replace_code)

	ReplaceCommon = replace_code
End Function

Function Map(Content)
	Set rs = Server.CreateObject("adodb.recordset")
	rs.open "select top 1 * from keyicms_Map", conn, 1, 2
	If not rs.eof Then
		MapX = rs("MapX")
		MapY = rs("MapY")
		MapLevel = rs("MapLevel")
		MapLogo = rs("MapLogo")
		AK = rs("AK")

		Content = replace(Content, "{keyicms:MapX}", MapX)
		Content = replace(Content, "{keyicms:MapY}", MapY)
		Content = replace(Content, "{keyicms:MapLevel}", MapLevel)
		Content = replace(Content, "{keyicms:MapLogo}", MapLogo)
		Content = replace(Content, "{keyicms:AK}", AK)
	End If
	Map = Content
	rs.close : set rs = nothing
End Function

Function NavList(Content, Ke01)
	NavList = Content
	labelRule = "{keyicms:Navigation"&str&"([\s\S]*?)}([\s\S]*?){/keyicms:Navigation"&str&"}"
	Set strDictionary = Server.CreateObject("scripting.dictionary")
	Set regExpObj = New Regexp
	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = labelRule
	set matches = regExpObj.Execute(Content)
	If matches.count > 0 Then
		For Each match in matches
			labelAttrLinklist = match.SubMatches(0)
			NavType = ParseArr(labelAttrLinklist)("NavType")
			NavList = replace(NavList, match.value, ParseNav("", match.value, Ke01, NavType))
			strDictionary.removeAll
		Next
	Else
		NavList = Content
	End If
	set matches = nothing
	Set regExpObj = nothing
End Function

Function ParseNav(str, Content, Ke01, NavType)
	If not IsExistStr(Content,"{keyicms:Navigation"&str&"") then Exit Function
	
	labelRule = "{keyicms:Navigation"&str&"([\s\S]*?)}([\s\S]*?){/keyicms:Navigation"&str&"}"
	labelRuleField = "{Navigation"&str&":([\s\S]+?)}"
	Set regExpObj = New Regexp
	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = labelRule
	set matches = regExpObj.Execute(Content)
	
	For Each match in matches
		labelAttrLinklist = match.SubMatches(0)
		ParentID = ParseArr(labelAttrLinklist)("ParentID") 	
		If IsNul(ParentID) then ParentID = 0
		lCurClass = ParseArr(labelAttrLinklist)("CurClass") 
			
		loopstrLinklist = match.SubMatches(1)
		Set rs = Server.CreateObject("adodb.recordset")
		If Ke01 = "PC" Then sql = "Select * from Keyicms_Navigation where ViewFlag and ParentID="&ParentID&" and NavType="&NavType&" Order by Sequence Asc, ID Desc"
		If Ke01 = "Mobile" Then sql = "Select * from Keyicms_Navigation where MobileFlag and ParentID="&ParentID&" and NavType="&NavType&" Order by Sequence Asc, ID Desc"
		rs.open sql, conn, 1, 2
		If not IsArray(linkArray) then vnum = -1 Else vnum = ubound(linkArray,2)
		regExpObj.IgnoreCase = True
		regExpObj.Global = True
		regExpObj.Pattern = labelRuleField
		set matchesfield = regExpObj.Execute(loopstrLinklist)
		loopstrTotal = ""
		If not rs.eof Then
			do until rs.eof
				If Ke01 = "PC" Then idCount = Conn.execute("Select Count(*) from Keyicms_Navigation where ViewFlag and ParentID="&rs("ID")&" and NavType="&NavType)(0)
				If Ke01 = "Mobile" Then idCount = Conn.execute("Select Count(*) from Keyicms_Navigation where MobileFlag and ParentID="&rs("ID")&" and NavType="&NavType)(0)
				loopstrLinklistNew = loopstrLinklist
				For Each matchfield in matchesfield
					fieldNameAndAttr = RegExpReplace(matchfield.SubMatches(0), "[\s]+", chr(32))
					m = instr(fieldNameAndAttr, chr(32))
					If m > 0 then 
						fieldName = left(fieldNameAndAttr, m - 1)
						fieldAttr = right(fieldNameAndAttr, len(fieldNameAndAttr) - m)
					Else
						fieldName = fieldNameAndAttr
						fieldAttr =	""
					End If
					Select Case fieldName
						Case "SubCount"
							loopstrLinklistNew = replace(loopstrLinklistNew, matchfield.value, idCount)
						Case "NavName"
							loopstrLinklistNew = replace(loopstrLinklistNew, matchfield.value, rs("NavName"))
						Case "Link"
							If rs("OutFlag") Then
								loopstrLinklistNew = replace(loopstrLinklistNew, matchfield.value, rs("OutUrl"))
							Else
								If Ke01 = "PC" Then loopstrLinklistNew = replace(loopstrLinklistNew, matchfield.value, rs("HtmlNavUrl"))
								If Ke01 = "Mobile" Then loopstrLinklistNew = replace(loopstrLinklistNew, matchfield.value, "/m"&rs("HtmlNavUrl"))
							End If						
						Case "ID"
							loopstrLinklistNew = replace(loopstrLinklistNew, matchfield.value, rs("ID"))	
					End Select
				Next
				loopstrTotal = loopstrTotal & loopstrLinklistNew
				rs.movenext
			loop
		End If
		set matchesfield = nothing
		Content = replace(Content, match.value, loopstrTotal)
	Next
	set matches=nothing
	If str = "" then str = 0
	If instr(Content,"{keyicms:Navigation"& str+1 )>0 then ParseNav str+1, Content, Ke01, NavType 'Else Exit Function
	ParseNav = ParseNav & Content
End Function

Function SortLoop(Content, Ke01)
	labelRule = "{keyicms:Sort([\s\S]*?)}([\s\S]*?){/keyicms:Sort}"
	Set regExpObj = New Regexp
	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = labelRule
	set matches = regExpObj.Execute(Content)
	If matches.count > 0 Then
		For Each match in matches
			labelAttrLinklist = match.SubMatches(0)
			SortType = ParseArr(labelAttrLinklist)("SortType")
			Content = replace(Content, match.value, ParseSort("", SortType, Ke01, match.value))
		Next
	Else
		Content = Content
	End If
	SortLoop = Content
	set matches = nothing
	Set regExpObj = nothing
End Function

Function ParseSort(str, SortType, Ke01, Content)
	If not IsExistStr(Content,"{keyicms:Sort"&str&"") then Exit Function
	Select Case SortType
		Case "About"
			DataFrom = "keyicms_About"
			Dir = AboutDir
		Case "News"
			DataFrom = "keyicms_NewsSort"
			Dir = NewsDir
		Case "Product"
			DataFrom = "keyicms_ProductSort"
			Dir = ProDir
		Case "Case"
			DataFrom = "keyicms_CaseSort"
			Dir = CaseDir
		Case "Other"
			DataFrom = "keyicms_OtherSort"
			Dir = OtherDir
		Case "Download"
			DataFrom = "keyicms_DownSort"
			Dir = DownDir
		Case Else
			ParseSort = Keyicms_Lang_ErrMsg(2)
			Exit Function
	End Select
	labelRule = "{keyicms:Sort"&str&"([\s\S]*?)}([\s\S]*?){/keyicms:Sort"&str&"}"
	labelRuleField = "{Sort"&str&":([\s\S]+?)}"
	Set regExpObj = New Regexp
	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = labelRule
	set matches = regExpObj.Execute(Content)
	
	If SortType = "About" Then
		For Each match in matches
			labelAttrLinklist = match.SubMatches(0)
			lCurClass = ParseArr(labelAttrLinklist)("CurClass")

			loopstrLinklist = match.SubMatches(1)
			Set rs = Server.CreateObject("adodb.recordset")
			If Ke01 = "PC" Then sql = "Select * from "&DataFrom&" where ViewFlag Order by Sequence Asc"
			If Ke01 = "Mobile" Then sql = "Select * from "&DataFrom&" where MobileFlag Order by Sequence Asc"
			rs.open sql, Conn, 1, 2

			Set regExpTwo = New Regexp
			regExpTwo.IgnoreCase = True
			regExpTwo.Global = True
			regExpTwo.Pattern = labelRuleField
			set matchesfield = regExpTwo.Execute(loopstrLinklist)
			loopstrTotal = ""
			If not rs.eof Then
				do until rs.eof
					loopstrLinklistNew = loopstrLinklist
					For Each matchfield in matchesfield
						fieldNameAndAttr = RegExpReplace(matchfield.SubMatches(0), "[\s]+", chr(32))
						m = instr(fieldNameAndAttr, chr(32))
						If m > 0 then 
							fieldName = left(fieldNameAndAttr, m - 1)
							fieldAttr = right(fieldNameAndAttr, len(fieldNameAndAttr) - m)
						Else
							fieldName = fieldNameAndAttr
							fieldAttr =	""
						End If
						Select Case fieldName
							Case "SortName"
								loopstrLinklistNew = replace(loopstrLinklistNew, matchfield.value, rs("AboutName"))
							Case "Link"
								If Ke01 = "PC" Then loopstrLinklistNew = replace(loopstrLinklistNew, matchfield.value, "/"&Dir&"/"&rs("ClassSeo")&Separated&rs("ID")&"."&HTMLName)
								If Ke01 = "Mobile" Then loopstrLinklistNew = replace(loopstrLinklistNew, matchfield.value, "/m/"&Dir&"/"&rs("ClassSeo")&Separated&rs("ID")&"."&HTMLName)
							Case "ID"
								loopstrLinklistNew = replace(loopstrLinklistNew, matchfield.value, rs("ID"))
							Case "AddTime"
								Format = ParseArr(fieldArr)("Format")
								nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("AddTime"), Format))
							Case "UpdateTime"
								Format = ParseArr(fieldArr)("Format")
								nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("UpdateTime"), Format))
							Case "Content"
								ContLen = ParseArr(fieldArr)("Len")
								If cInt(ContLen) > 0 Then
									ContLen = cInt(ContLen)
									Cont = StrLeft(RemoveHTML(rs("Content")), ContLen)
								Else
									Cont = RemoveHTML(rs("Content"))
								End If
								nloopstr = replace(nloopstr, matchfield.value, Cont)
						End Select
					Next
					loopstrTotal = loopstrTotal & loopstrLinklistNew
					rs.movenext
				loop
			End If
			set matchesfield = nothing
			set regExpTwo = nothing
			Content = replace(Content, match.value, loopstrTotal)
		Next
		set matches = nothing
		ParseSort = ParseSort & Content
	Else
		For Each match in matches
			labelAttrLinklist = match.SubMatches(0)
			loopstrLinklist = match.SubMatches(1)
			ParentID = ParseArr(labelAttrLinklist)("ParentID") 	
			If IsNul(ParentID) then ParentID = 0
	
			Set rs = Server.CreateObject("adodb.recordset")
			rs.open "Select * from "&DataFrom&" where ViewFlag and ParentID="&ParentID, Conn, 1, 2
			Set regExpTwo = New Regexp
			regExpTwo.IgnoreCase = True
			regExpTwo.Global = True
			regExpTwo.Pattern = labelRuleField
			set matchesfield = regExpTwo.Execute(loopstrLinklist)
			loopstrTotal = ""
			If not rs.eof Then
				do until rs.eof
					idCount = Conn.execute("Select Count(*) from "&DataFrom&" where ViewFlag and ParentID="&rs("ID"))(0)
					loopstrLinklistNew = loopstrLinklist
					For Each matchfield in matchesfield
						fieldNameAndAttr = RegExpReplace(matchfield.SubMatches(0), "[\s]+", chr(32))
						m = instr(fieldNameAndAttr, chr(32))
						If m > 0 then 
							fieldName = left(fieldNameAndAttr, m - 1)
							fieldAttr = right(fieldNameAndAttr, len(fieldNameAndAttr) - m)
						Else
							fieldName = fieldNameAndAttr
							fieldAttr =	""
						End If
						Select Case fieldName
							Case "SubCount"
								loopstrLinklistNew = replace(loopstrLinklistNew, matchfield.value, idCount)
							Case "SortName"
								loopstrLinklistNew = replace(loopstrLinklistNew, matchfield.value, rs("SortName"))
							Case "SmallPic"
								loopstrLinklistNew = replace(loopstrLinklistNew, matchfield.value, rs("SmallPic"))
							Case "Content"
								loopstrLinklistNew = replace(loopstrLinklistNew, matchfield.value, rs("Content"))
							Case "Link"
								If Ke01 = "PC" Then loopstrLinklistNew = replace(loopstrLinklistNew, matchfield.value, "/"&Dir&"/"&rs("FolderName"))
								If Ke01 = "Mobile" Then loopstrLinklistNew = replace(loopstrLinklistNew, matchfield.value, "/m/"&Dir&"/"&rs("FolderName"))
							Case "ID"
								loopstrLinklistNew = replace(loopstrLinklistNew, matchfield.value, rs("ID"))						
						End Select
					Next
					loopstrTotal = loopstrTotal & loopstrLinklistNew
					rs.movenext
				loop
			End If
			set matchesfield = nothing
			set regExpTwo = nothing
			Content = replace(Content, match.value, loopstrTotal)
		Next
		set matches = nothing
		If str = "" then str = 0
		If instr(Content, "{keyicms:Sort"& str+1 )>0 then ParseSort str+1, SortType, Ke01, Content 'Else Exit Function
		ParseSort = ParseSort & Content
	End If
End Function

Public Function KE01_If(Content)
	labelRule = "\[If:([\s\S]+?)\]([\s\S]*?)\[End\s+If"&x&"\]"
	Set regExpObj = New Regexp
	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = labelRule
	set matches = regExpObj.Execute(Content)
	If matches.count > 0 Then
		For Each match in matches
			Content = replace(Content, match.value, ParseIf(Content, ""))
		Next
	Else
		Content = Content
	End If
	KE01_If = Content
	set matches = nothing
	Set regExpObj = nothing
End Function

Public Function ParseIf(Content, x)
	If not IsExistStr(Content, "[If"&x&":") then Exit Function
	Dim nextone
	If IsNum(x) then
		nextone = Int(x) + 1
	Else
		nextone = 1
	End If
	If IsExistStr(Content,"[If"&nextone&":") then ParseIf(nextone)
	Dim matchIf, matchesIf, strIf, strThen, strThen1, strElse1, labelRule2, labelRule3
	Dim IfFlag, ElseIfArray, ElseIfSubArray, ElseIfArrayLen, resultStr, ElseIfLen, strElseIf, strElseIfThen, ElseIfFlag
	labelRule = "\[If"&x&":([\s\S]+?)\]([\s\S]*?)\[End\s+If"&x&"\]" : labelRule2 = "[ElseIf"&x&"":labelRule3="[Else]" : ElseIfFlag = false
	Set regExpObj = New Regexp
	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = labelRule

	set matchesIf = regExpObj.Execute(Content)
	For Each matchIf in matchesIf 			
		strIf = matchIf.SubMatches(0) : strThen = matchIf.SubMatches(1)
		If instr(strThen, labelRule2) > 0 then
			ElseIfArray = split(strThen, labelRule2) : ElseIfArrayLen = ubound(ElseIfArray) : ElseIfSubArray = split(ElseIfArray(ElseIfArrayLen), labelRule3)
			If instr(ElseIfArray(ElseIfArrayLen), labelRule3) > 0 then
				resultStr = ElseIfSubArray(1)
			End If
			Execute("If "&strIf&" then resultStr=ElseIfArray(0)")
			for ElseIfLen = 1 to ElseIfArrayLen - 1
				strElseIf = GetSubStrByFromAndEnd(ElseIfArray(ElseIfLen),":","}","")
				strElseIfThen = GetSubStrByFromAndEnd(ElseIfArray(ElseIfLen),"}","","start")
				Execute("If "&strElseIf&" then resultStr = strElseIfThen")
				Execute("If "&strElseIf&" then ElseIfFlag = true Else ElseIfFlag = false")
				If ElseIfFlag then exit for
			Next
			Execute("If "&GetSubStrByFromAndEnd(ElseIfSubArray(0),":","}","")&" then resultStr=GetSubStrByFromAndEnd(ElseIfSubArray(0),""}"","""",""start""):ElseIfFlag=true")
			Content = replace(Content, matchIf.value, resultStr)
		Else 
			If instr(strThen, labelRule3) > 0 then 
				strThen1 = split(strThen, labelRule3)(0)
				strElse1 = split(strThen, labelRule3)(1)

				Execute("If "&strIf&" then IfFlag = true Else IfFlag = false")
				If IfFlag then Content = replace(Content, matchIf.value, strThen1) Else Content = replace(Content, matchIf.value, strElse1)
			Else
				Execute("If "&strIf&" then IfFlag = true Else IfFlag = false")
				If IfFlag then Content = replace(Content, matchIf.value, strThen) Else Content = replace(Content, matchIf.value, "")
			End If
		End If
		elseIfFlag = false
	Next
	set matchesIf = nothing
	ParseIf = Content
End Function

Function ShowPage(replace_code, idcount, Page, Pagec, Pages)
	Dim Pagenc, Pagenmax, Pagenmin
	Pagenc = 2

	Dim Myself, PATH_INFO, QUERY_STRING
	PATH_INFO = request.servervariables("PATH_INFO")
	QUERY_STRING = request.ServerVariables("QUERY_STRING")'
	Myself = "Index"

	Pagenmin = Page - Pagenc
	Pagenmax = Page + Pagenc
	If(Pagenmin < 1) Then Pagenmin = 1
	If(Pagenmax > Pagec) Then Pagenmax = Pagec

	Set strDictionary = Server.CreateObject("scripting.dictionary")
	Set regExpObj = New Regexp
	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = "{keyicms:Page([\s\S]*?)}([\s\S]*?){/keyicms:Page}"
	set matches = regExpObj.Execute(replace_code)
	For Each match in matches
		If idcount > 0 Then
			labelStr = match.SubMatches(0)
			loopStr = match.SubMatches(1)
			
			Set regExpObjTow = New Regexp
			regExpObjTow.IgnoreCase = True
			regExpObjTow.Global = True
			regExpObjTow.Pattern = "{Page:PageList([\s\S]*?)}([\s\S]*?){/Page:PageList}"
			set matchesTow = regExpObjTow.Execute(loopStr)
			For Each matchTow in matchesTow
				labelStrTow = matchTow.SubMatches(0)
				set labelArrTow = ParseArr(labelStrTow)
				lCurrent = labelArrTow("PageCur")

				loopStrTow = matchTow.SubMatches(1)
	
				Set regExpObjNew = New Regexp
				regExpObjNew.IgnoreCase = True
				regExpObjNew.Global = True
				regExpObjNew.Pattern = "{PageList:([\s\S]*?)}"
				set matchesfield = regExpObjNew.Execute(loopStrTow)
	
				loopstrTotal = ""
				For j = Pagenmin to Pagenmax
					nloopstr = loopStrTow
					For Each matchfield in matchesfield
						fieldNameArr = RegExpReplace(matchfield.SubMatches(0),"[\s]+",chr(32))
						fieldNameArr = trim(fieldNameArr)
						m = instr(fieldNameArr,chr(32))
						If  m > 0 then 
							fieldArr=""
							fieldName = left(fieldNameArr,m - 1)
							fieldArr =	right(fieldNameArr,len(fieldNameArr) - m)
						Else
							fieldName = fieldNameArr
							fieldArr =	""
						End If
		
						Select Case fieldName
							Case "Link"
								nloopstr = replace(nloopstr, matchfield.value, Myself & Separated & j &"."&HTMLName)
							Case "PageInfo"
								nloopstr = replace(nloopstr, matchfield.value, j)
							Case "PageCur"
								If Page = j Then
									nloopstr = replace(nloopstr, matchfield.value, lCurrent)
								Else
									nloopstr = replace(nloopstr, matchfield.value, "")
								End If
						End Select
					Next
					loopstrTotal = loopstrTotal & nloopstr
				Next
				loopStr = replace(loopStr, matchTow.value, loopstrTotal)
			Next
	
			Set regExpObjPage = New Regexp
			regExpObjPage.IgnoreCase = True
			regExpObjPage.Global = True
			regExpObjPage.Pattern = "{Page:([\s\S]*?)}([\s\S]*?){/Page:([\s\S]*?)}"
			set matchesPage = regExpObjPage.Execute(loopStr)
			nloopstr = loopStr
			For Each matchPage in matchesPage
				labelStr = matchPage.SubMatches(0)
				loopStr = matchPage.SubMatches(1)
				Select Case labelStr
					Case "First"
						nloopstr = replace(nloopstr, matchPage.value, "<a href='"&Myself & Separated & "1."&HTMLName&"'>"&loopStr&"</a>")
					Case "Prev"
						If Page > 1 Then
							nloopstr = replace(nloopstr, matchPage.value, "<a href='"&Myself & Separated & (Page-1) & "."&HTMLName&"'>"&loopStr&"</a>")
						Else
							nloopstr = replace(nloopstr, matchPage.value, "")
						End If
					Case "Next"
						If Page < Pagec Then
							nloopstr = replace(nloopstr, matchPage.value, "<a href='"&Myself & Separated & (Page+1) & "."&HTMLName&"'>"&loopStr&"</a>")
						Else
							nloopstr = replace(nloopstr, matchPage.value, "")
						End If
					Case "Last"
						nloopstr = replace(nloopstr, matchPage.value, "<a href='"&Myself & Separated & Pagec & "."&HTMLName&"'>"&loopStr&"</a>")
				End Select
			Next
			replace_code = replace(replace_code, match.value, nloopstr)
		Else
			replace_code = replace(replace_code, match.value, "")
		End If
	Next
	ShowPage = ShowPage & replace_code
End Function

'搜索 __ 翻页
Function ShowPageAsp(replace_code, idcount, Page, Pagec, Pages)
	Dim Pagenc, Pagenmax, Pagenmin
	Pagenc = 2

	Dim Myself, PATH_INFO, QUERY_STRING
	PATH_INFO = request.servervariables("PATH_INFO")
	QUERY_STRING = request.ServerVariables("QUERY_STRING")'
	
	If QUERY_STRING = "" Or InStr(PATH_INFO & "?" & QUERY_STRING, "Page=") = 0 Then
		Myself = PATH_INFO & "?" & QUERY_STRING
	Else
		Myself = Left(PATH_INFO & "?" & QUERY_STRING, InStr(PATH_INFO & "?" & QUERY_STRING, "Page=")-2)
	End If

	Pagenmin = Page - Pagenc
	Pagenmax = Page + Pagenc
	If(Pagenmin < 1) Then Pagenmin = 1
	If(Pagenmax > Pagec) Then Pagenmax = Pagec

	Set strDictionary = Server.CreateObject("scripting.dictionary")
	Set regExpObj = New Regexp
	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = "{keyicms:Page([\s\S]*?)}([\s\S]*?){/keyicms:Page}"
	set matches = regExpObj.Execute(replace_code)
	For Each match in matches
		If idcount > 0 Then
			labelStr = match.SubMatches(0)
			loopStr = match.SubMatches(1)
			
			Set regExpObjTow = New Regexp
			regExpObjTow.IgnoreCase = True
			regExpObjTow.Global = True
			regExpObjTow.Pattern = "{Page:PageList([\s\S]*?)}([\s\S]*?){/Page:PageList}"
			set matchesTow = regExpObjTow.Execute(loopStr)
			For Each matchTow in matchesTow
				labelStrTow = matchTow.SubMatches(0)
				set labelArrTow = ParseArr(labelStrTow)
				lCurrent = labelArrTow("PageCur")
				loopStrTow = matchTow.SubMatches(1)
	
				Set regExpObjNew = New Regexp
				regExpObjNew.IgnoreCase = True
				regExpObjNew.Global = True
				regExpObjNew.Pattern = "{PageList:([\s\S]*?)}"
				set matchesfield = regExpObjNew.Execute(loopStrTow)
	
				loopstrTotal = ""
				For j = Pagenmin to Pagenmax
					nloopstr = loopStrTow
					For Each matchfield in matchesfield
						fieldNameArr = RegExpReplace(matchfield.SubMatches(0),"[\s]+",chr(32))
						fieldNameArr = trim(fieldNameArr)
						m = instr(fieldNameArr,chr(32))
						If  m > 0 then 
							fieldArr=""
							fieldName = left(fieldNameArr,m - 1)
							fieldArr =	right(fieldNameArr,len(fieldNameArr) - m)
						Else
							fieldName = fieldNameArr
							fieldArr =	""
						End If
		
						Select Case fieldName
							Case "Link"
								nloopstr = replace(nloopstr, matchfield.value, Myself & "&Page=" & j)
							Case "PageInfo"
								nloopstr = replace(nloopstr, matchfield.value, j)
							Case "PageCur"
								If Page = j Then
									nloopstr = replace(nloopstr, matchfield.value, lCurrent)
								Else
									nloopstr = replace(nloopstr, matchfield.value, "")
								End If
						End Select
					Next
					loopstrTotal = loopstrTotal & nloopstr
				Next
				loopStr = replace(loopStr, matchTow.value, loopstrTotal)
			Next
	
			Set regExpObjPage = New Regexp
			regExpObjPage.IgnoreCase = True
			regExpObjPage.Global = True
			regExpObjPage.Pattern = "{Page:([\s\S]*?)}([\s\S]*?){/Page:([\s\S]*?)}"
			set matchesPage = regExpObjPage.Execute(loopStr)
			nloopstr = loopStr
			For Each matchPage in matchesPage
				labelStr = matchPage.SubMatches(0)
				loopStr = matchPage.SubMatches(1)
				Select Case labelStr
					Case "First"
						nloopstr = replace(nloopstr, matchPage.value, "<a href='"&Myself & "&Page=1'>"&loopStr&"</a>")
					Case "Prev"
						If Page > 1 Then
							nloopstr = replace(nloopstr, matchPage.value, "<a href='"&Myself & "&Page=" & (Page-1) &"'>"&loopStr&"</a>")
						Else
							nloopstr = replace(nloopstr, matchPage.value, "")
						End If
					Case "Next"
						If Page < Pagec Then
							nloopstr = replace(nloopstr, matchPage.value, "<a href='"&Myself & "&Page=" & (Page+1) & "'>"&loopStr&"</a>")
						Else
							nloopstr = replace(nloopstr, matchPage.value, "")
						End If
					Case "Last"
						nloopstr = replace(nloopstr, matchPage.value, "<a href='"&Myself & "&Page=" & Pagec & "'>"&loopStr&"</a>")
				End Select
			Next
			replace_code = replace(replace_code, match.value, nloopstr)
		Else
			replace_code = replace(replace_code, match.value, "")
		End If
	Next
	ShowPageAsp = ShowPageAsp & replace_code
End Function

Function ParseTime(Content, tValue, tSort)
	set regExpTwo = New Regexp
	regExpTwo.IgnoreCase = True
	regExpTwo.Global = True
	regExpTwo.Pattern = "{keyicms:"&tSort&"([\s\S]*?)}"
	set matchesfield = regExpTwo.Execute(Content)
	For Each matchfield in matchesfield
		fieldNameArr = RegExpReplace(matchfield.SubMatches(0),"[\s]+",chr(32))
		fieldNameArr = trim(fieldNameArr)
		m = instr(fieldNameArr,chr(32))
		If  m > 0 then 
			fieldArr = ""
			fieldName = left(fieldNameArr,m - 1)
			fieldArr =	right(fieldNameArr,len(fieldNameArr) - m)
		Else
			fieldName = fieldNameArr
			fieldArr =	""
		End If
		Format = ParseArr(fieldName)("Format")
		If IsNul(Format) Then
			Content = replace(Content, matchfield.value, tValue)
		Else
			Content = replace(Content, matchfield.value, FormatDate(tValue, Format))
		End If
	Next
	ParseTime = Content
End Function

Function Position(Content, SortID, DataForm, Ke01)
	If not IsExistStr(Content, "{keyicms:Position}") or not IsExistStr(Content, "{/keyicms:Position}") then
		Position = Content
		Exit Function
	End If
	If DataForm = 1 Then FormName = "keyicms_About"
	If DataForm = 2 Then FormName = "keyicms_NewsSort"
	If DataForm = 3 Then FormName = "keyicms_ProductSort"
	If DataForm = 4 Then FormName = "keyicms_CaseSort"
	If DataForm = 5 Then FormName = "keyicms_DownSort"
	If DataForm = 6 Then FormName = "keyicms_JobSort"

	Set regExpObj = New Regexp
	set strDictionary = server.CreateObject("scripting.dictionary")
	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = "{keyicms:Position}([\s\S]*?){/keyicms:Position}"
	set matches = regExpObj.Execute(Content)

  	for each match in matches
		loopStr = match.SubMatches(0)
		loopstrTotal = ""
		Content = replace(Content, match.value, KE01_Position(loopStr, SortID, DataForm, FormName))
		loopstrTotal = loopstrTotal & nloopstr
		strDictionary.removeAll
	next
	Position = Content
End Function

Function KE01_Position(loopStr, SortID, DataForm, FormName)
    Dim rs, sql
	Set rs = server.CreateObject("adodb.recordset")
	sql = "select * from keyicms_System where DataForm="&DataForm
	rs.open sql, Conn, 1, 2
	If Ke01 = "PC" Then AutoLink = "/" & rs("SystemDir")
	If Ke01 = "Mobile" Then AutoLink = "/m/" & rs("SystemDir")
	SortName = rs("SeoTitle")
	rs.close : set rs = nothing

	If SortID <> "" Then
		Set rs = server.CreateObject("adodb.recordset")
		sql = "Select * From "&FormName&" where ViewFlag and ID="&SortID
		rs.Open sql, conn, 1, 1
		If Not rs.EOF Then
			SortName = rs("SortName")
			AutoLink = AutoLink&"/"&rs("FolderName")
			KE01_Position = KE01_Position(loopStr, rs("ParentID"), DataForm, FormName)
		End If
		rs.close : set rs = nothing
	End If

	nloopstr = loopStr
	set regExpTwo = New Regexp
	regExpTwo.IgnoreCase = True
	regExpTwo.Global = True
	regExpTwo.Pattern = "{Position:([\s\S]*?)}"
	set matchesfield = regExpTwo.Execute(loopStr)
	
	for each matchfield in matchesfield
		FieldNameArr = RegExpReplace(matchfield.SubMatches(0), "[\s]+", chr(32))
		FieldNameArr = trim(FieldNameArr)
		m = instr(FieldNameArr, chr(32))
		if  m > 0 then 
			FieldName = left(FieldNameArr, m - 1)
			FieldArr =	right(FieldNameArr, len(FieldNameArr) - m)
		else
			FieldName = FieldNameArr
			FieldArr =	""
		end if
		Select Case FieldName
			Case "Link"
				nloopstr = replace(nloopstr, matchfield.value, AutoLink)
			Case "PosName"
				nloopstr = replace(nloopstr, matchfield.value, SortName)
		end select
	next
	loopstrTotal = loopstrTotal & nloopstr
	set matchesfield = nothing
	KE01_Position = KE01_Position & loopstrTotal
End Function

'获取可用标签参数
Function ParseArr(Byval attr)
	Set strDictionary = Server.CreateObject("scripting.dictionary")
	Dim attrStr, attrArray, attrDictionary, i, singleAttr, singleAttrKey, singleAttrValue
	attrStr = RegExpReplace(attr, "[\s]+", chr(32))
	attrStr = trim(attrStr)
	attrArray = split(attrStr, chr(32))
	for i = 0 to ubound(attrArray)
		singleAttr = split(attrArray(i), chr(61))
		singleAttrKey = singleAttr(0)
		singleAttrValue = singleAttr(1)
		
		If not strDictionary.Exists(singleAttrKey) then
			strDictionary.add singleAttrKey, singleAttrValue
		Else
			strDictionary(singleAttrKey) = singleAttrValue
		End If
	Next
	set ParseArr = strDictionary
End Function

Function RegExpReplace(ContentStr, PatternStr, ReplaceStr)
	Set regExpObj = New Regexp
	regExpObj.Pattern = PatternStr
	RegExpReplace = regExpObj.replace(ContentStr, ReplaceStr)
End Function

Function IsNul(str)
	If IsNull(str) or str = "" then IsNul = true Else IsNul = false
End Function

'在str中是否存在findstr
Function IsExistStr(Content, findstr)
	If IsNul(Content) or IsNul(findstr) then IsExistStr = false : Exit Function
	If instr(Content, findstr)>0 then IsExistStr = true Else IsExistStr = false
End Function

'是否为数字
Function IsNum(str)
	If not IsNul(str) then IsNum = IsNumeric(str) Else IsNum=false
End Function

Function ChangeNull(str)
	If IsNull(str) Then
		ChangeNull = ""
	Else
		ChangeNull = str
	End If
End Function

Function GetSubStrByFromAndEnd(Str, StartStr, EndStr, OperType)
	Dim location1, location2
	Select Case OperType
		Case "start"
			location1 = instr(Str, StartStr) + len(StartStr) : location2 = len(Str) + 1
		Case "end"
			location1 = 1 : location2=instr(location1,str,endStr)
		Case Else
			location1 = instr(Str, StartStr) + len(StartStr) : location2 = instr(location1, Str, EndStr)
	End Select
	GetSubStrByFromAndEnd = mid(str, location1, location2 - location1) 
End Function
%>