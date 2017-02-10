<!--#Include file="html_Common.Asp"-->
<!--#Include file="html_To.Asp"-->
<%
'生成静态Function方法
dim ExplainJS, ConstJs

ExplainJS = "<script>window.location.replace('/Maintain.html')</script>"
If mStatus and MobileJump Then
	ConstJs = ConstJs & "<script>if(navigator.platform.indexOf('Win32')<0){window.location.replace('/m'+window.location.pathname)}</script>"
End If
If EnableCopy Then
	ConstJs = ConstJs & "<style type='text/css' media='screen'>body{ -moz-user-select:none; -webkit-user-select:none;}</style>" & vbCrLf
	ConstJs = ConstJs & "<script type='text/javascript'>" & vbCrLf
	ConstJs = ConstJs & "document.onselectstart=function(e){return false;}" & vbCrLf
	ConstJs = ConstJs & "document.oncontextmenu=function(e){return false;}" & vbCrLf
	ConstJs = ConstJs & "</script>" & vbCrLf
End If

Function ReplaceLoop(Label, Content)
	Set regExpObj = New Regexp
	set strDictionary = server.CreateObject("scripting.dictionary")

	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = "{keyicms:"&Label&"([\s\S]*?)}([\s\S]*?){/keyicms:"&Label&"}"
	
	set matches = regExpObj.Execute(Content)
  	for each match in matches
		labelStr = match.SubMatches(0)
		loopStr = match.SubMatches(1)

		If Label = "FriendLink" Then
			ReplaceStr = FriendLink(labelStr, loopStr)
		End If
		If Label = "Slide" Then
			ReplaceStr = Slide(loopStr)
		End If
		If Label = "Ads" Then
			ReplaceStr = Ads(labelStr, loopStr)
		End If
		If Label = "Contact" Then
			ReplaceStr = Contact(labelStr, loopStr)
		End If
		Content = replace(Content, match.value, ReplaceStr)
		strDictionary.removeAll
	next
	set matches = nothing
	ReplaceLoop = Content
End Function

Function Include(Content, TemplatePath)
	Set regExpObj = New Regexp
	set strDictionary = server.CreateObject("scripting.dictionary")

	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = "{keyicms:Include([\s\S]*?)}"

	set matches = regExpObj.Execute(Content)
  	for each match in matches
		labelStr = match.SubMatches(0)

		set labelArr = ParseArr(labelStr)
		lFile = labelArr("File")
		set matchesfield = nothing
		Content = replace(Content, match.value,  ReadFromUTF(TemplatePath&"/html/"&lFile, "utf-8"))
		strDictionary.removeAll
	next
	set matches = nothing
	Include = Content
End Function

Function About(Content, Page, Ke01, SortPath)
	Set regExpObj = New Regexp
	set strDictionary = server.CreateObject("scripting.dictionary")

	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = "{keyicms:About([\s\S]*?)}([\s\S]*?){/keyicms:About}"

	set matches = regExpObj.Execute(Content)
  	for each match in matches
		labelStr = match.SubMatches(0)
		loopStr = match.SubMatches(1)

		set labelArr = ParseArr(labelStr)
		lID = labelArr("ID")
		lOrder = labelArr("Order")

		If Ke01 = "PC" Then datawhere = " where ViewFlag"
		If Ke01 = "Mobile" Then datawhere = " where MobileFlag"

		If not IsNul(lID) Then
			datawhere = datawhere & " and ID="&lID
		End If

		datafrom = "keyicms_About"
		Set regExpTwo = New Regexp
		regExpTwo.IgnoreCase = True
		regExpTwo.Global = True
		regExpTwo.Pattern = "{About:([\s\S]*?)}"
		set matchesfield = regExpTwo.Execute(loopStr)
	
		Set rs = server.CreateObject("adodb.recordset")
		sql = "select * from ["& datafrom &"]" & datawhere & taxis
		rs.open sql, Conn, 1, 2
		If rs.eof Then
			loopstrTotal = Keyicms_Lang_ErrMsg(3)
		Else
			loopstrTotal = ""
			do until rs.eof
				If Ke01 = "PC" Then AutoLink = "/"&AboutDir&"/"&rs("ClassSeo")&Separated&rs("ID")&"."&HTMLName
				If Ke01 = "Mobile" Then AutoLink = "/m/"&AboutDir&"/"&rs("ClassSeo")&Separated&rs("ID")&"."&HTMLName
				nloopstr = loopStr
				for each matchfield in matchesfield
					FieldNameArr = regExpReplace(matchfield.SubMatches(0), "[\s]+", chr(32))
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
						Case "ID"
							nloopstr = replace(nloopstr, matchfield.value, rs("ID"))
						Case "AboutName"
							NameLen = ParseArr(FieldArr)("Len") 
							If cInt(NameLen) > 0 Then
								NameLen = cInt(NameLen)
								Title = StrLeft(rs("AboutName"), NameLen)
							Else
								Title = rs("AboutName")
							End If
							nloopstr = replace(nloopstr, matchfield.value, Title)
						Case "AddTime"
							Format = ParseArr(FieldArr)("Format")
							nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("AddTime"), Format))
						Case "UpdateTime"
							Format = ParseArr(FieldArr)("Format")
							nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("UpdateTime"), Format))
						Case "Content"
							ContLen = ParseArr(FieldArr)("Len")
							If cInt(ContLen) > 0 Then
								ContLen = cInt(ContLen)
								Cont = StrLeft(RemoveHTML(rs("Content")), ContLen)
							Else
								Cont = RemoveHTML(rs("Content"))
							End If
							nloopstr = replace(nloopstr, matchfield.value, Cont)
					end select
				next
				loopstrTotal = loopstrTotal & nloopstr
				rs.movenext
			loop
			rs.close : set rs = nothing
		End If
		set matchesfield = nothing
		Content = replace(Content, match.value, loopstrTotal)
		strDictionary.removeAll
	next
	About = Content
End Function

Function NewsList(Content, Page, Ke01, SortPath)
	If Page = "" Then Page = 1
	Set regExpObj = New Regexp
	set strDictionary = server.CreateObject("scripting.dictionary")

	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = "{keyicms:NewsList([\s\S]*?)}([\s\S]*?){/keyicms:NewsList}"

	set matches = regExpObj.Execute(Content)
  	for each match in matches
		labelStr = match.SubMatches(0)
		loopStr = match.SubMatches(1)

		set labelArr = ParseArr(labelStr)
		lNum = labelArr("Num")

		lSortID = labelArr("SortID")
		lSortPath = labelArr("SortPath")
		If lSortID = "" and lSortPath = "" Then
			lSortPath = "0,"
		ElseIf lSortID = "" and lSortPath <> "" Then
			lSortPath = lSortPath
		ElseIf lSortID <> "" Then
			sql = "select * from keyicms_NewsSort where ID="&lSortID
			If Conn.execute(sql).Eof Then
				NewsList = "SortID"&Keyicms_Lang_ErrMsg(0)
				Exit Function
			Else
				lSortPath = Conn.execute(sql)("SortPath")
			End If
		End If

		lPicFlag = labelArr("PicFlag")
		lNoticeFlag = labelArr("NoticeFlag")
		lOrder = labelArr("Order")
		lPageFlag = labelArr("PageFlag")

		If lPageFlag Then
			If IsNul(SortPath) Then
				dataSortPath = ""
			Else
				dataSortPath = " and Instr(SortPath,'"&SortPath&"')>0"
			End If
		Else
			If IsNul(lSortPath) Then
				dataSortPath = ""
			Else
				dataSortPath = " and Instr(SortPath,'"&lSortPath&"')>0"
			End If
		End If
		
		dim HideSort
		Set rs = server.CreateObject("adodb.recordset")
		rs.open "select * from [keyicms_NewsSort] Where not ViewFlag" & dataSortPath, Conn, 1, 2
		while not rs.eof
			HideSort = "and not(Instr(SortPath,'"&rs("SortPath")&"')>0) " & HideSort
			rs.movenext
		wend
		rs.close : set rs = nothing

		If Ke01 = "PC" Then datawhere = " where ViewFlag"
		If Ke01 = "Mobile" Then datawhere = " where MobileFlag"
		
		datawhere = datawhere & dataSortPath
		If lPicFlag = "True" Then
			datawhere = datawhere & " and PicFlag"
		ElseIf lPicFlag = "False" Then
			datawhere = datawhere & " and not PicFlag"
		End IF
		If lNoticeFlag = "True" Then
			datawhere = datawhere & " and NoticeFlag"
		ElseIf lNoticeFlag = "False" Then
			datawhere = datawhere & " and not NoticeFlag"
		End IF
	
		If lOrder="AddTime" Then
			taxis = " Order by "&lOrder&" Desc"
		ElseIf lOrder="Notice" Then
			taxis = " Order by NoticeFlag Asc, ID Desc"
		Else
			taxis = " Order by ID Desc"
		End If
		
		datafrom = "keyicms_News"
		Pages = lNum

		Dim rs, sql
		Set rs = server.CreateObject("adodb.recordset")
		sql = "select * from ["& datafrom &"]" & datawhere & HideSort
		rs.Open sql, Conn, 1, 2
		idCount = rs.recordcount
		rs.close : set rs = nothing

		If idCount > 0 Then
			If(idCount Mod Pages = 0)Then
				Pagec = Int(idCount / Pages)
			Else
				Pagec = Int(idCount / Pages) + 1
			End If
			Set rs = server.CreateObject("adodb.recordset")
			sql = "select * from ["& datafrom &"]" & datawhere & HideSort & taxis
			rs.Open sql, Conn, 1, 1
			rs.Pagesize = Pages
			If Page < 1 Then Page = 1
			If lPageFlag = "True" Then
				If Pagec > 0 Then rs.absolutePage = Page
			Else
				If Pagec > 0 Then rs.absolutePage = 1
			End If
			For j = 1 To rs.Pagesize
				If rs.EOF Then Exit For
				If(j = 1)Then
					sqlid = rs("ID")
				Else
					sqlid = sqlid &","&rs("ID")
				End If
				rs.movenext
			Next
			rs.close : set rs = nothing
		End If
			
		Set regExpTwo = New Regexp
		regExpTwo.IgnoreCase = True
		regExpTwo.Global = True
		regExpTwo.Pattern = "{NewsList:([\s\S]*?)}"
		set matchesfield = regExpTwo.Execute(loopStr)

		If idCount > 0 And sqlid <> "" Then
			Set rs = server.CreateObject("adodb.recordset")
			sql = "select * from ["& datafrom &"] where ID in("& sqlid &") " & taxis
			rs.open sql, Conn, 1, 2
			loopstrTotal = ""
			For i = 1 To rs.recordcount
				If Ke01 = "PC" Then AutoLink = "/"&NewsDir&"/"&rs("ClassSeo")&Separated&rs("ID")&"."&HTMLName
				If Ke01 = "Mobile" Then AutoLink = "/m/"&NewsDir&"/"&rs("ClassSeo")&Separated&rs("ID")&"."&HTMLName
				nloopstr = loopStr
				for each matchfield in matchesfield
					FieldNameArr = regExpReplace(matchfield.SubMatches(0), "[\s]+", chr(32))
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
						Case "i"
							nloopstr = replace(nloopstr, matchfield.value, i)
						Case "Link"
							nloopstr = replace(nloopstr, matchfield.value, AutoLink)
						Case "SortName"
							SortName = Conn.execute("Select * from keyicms_NewsSort where ID="&rs("SortID"))("SortName")
							nloopstr = replace(nloopstr, matchfield.value, SortName)
						Case "SortID"
							nloopstr = replace(nloopstr, matchfield.value, rs("SortID"))
						Case "ID"
							nloopstr = replace(nloopstr, matchfield.value, rs("ID"))
						Case "NewsName"
							NameLen = ParseArr(FieldArr)("Len") 
							If cInt(NameLen) > 0 Then
								NameLen = cInt(NameLen)
								Title = StrLeft(rs("NewsName"), NameLen)
							Else
								Title = rs("NewsName")
							End If
							nloopstr = replace(nloopstr, matchfield.value, Title)
						Case "Origin"
							nloopstr = replace(nloopstr, matchfield.value, changeNull(rs("Origin")))
						Case "SmallPic"
							nloopstr = replace(nloopstr, matchfield.value, changeNull(rs("SmallPic")))
						Case "Color"
							nloopstr = replace(nloopstr, matchfield.value, rs("Color"))
						Case "Bold"
							nloopstr = replace(nloopstr, matchfield.value, rs("Bold"))
						Case "Text"
							TextLen = ParseArr(FieldArr)("Len")
							TextLen = cInt(TextLen)
							nloopstr = replace(nloopstr, matchfield.value, StrLeft(rs("Text"), TextLen))
						Case "AddTime"
							Format = ParseArr(FieldArr)("Format")
							nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("AddTime"), Format))
						Case "UpdateTime"
							Format = ParseArr(FieldArr)("Format")
							nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("UpdateTime"), Format))
						Case "ClickNumber"
							nloopstr = replace(nloopstr, matchfield.value, rs("ClickNumber"))
						Case "Content"
							ContLen = ParseArr(FieldArr)("Len")
							If cInt(ContLen) > 0 Then
								ContLen = cInt(ContLen)
								Cont = StrLeft(RemoveHTML(rs("Content")), ContLen)
							Else
								Cont = RemoveHTML(rs("Content"))
							End If
							nloopstr = replace(nloopstr, matchfield.value, Cont)
					end select
				next
				loopstrTotal = loopstrTotal & nloopstr
				rs.movenext
			Next
			rs.close : set rs = nothing
		Else
			loopstrTotal = Keyicms_Lang_ErrMsg(1)
		End If
		set matchesfield = nothing
		Content = replace(Content, match.value, loopstrTotal)
		strDictionary.removeAll
		If lPageFlag = "True" Then
			Content = ShowPage(Content, idCount, Page, Pagec, Pages)
		End If
	next
	set matches = nothing
	NewsList = Content
End Function

Function ProductList(Content, Page, Ke01, SortPath)
	If Page = "" Then Page = 1
	Set regExpObj = New Regexp
	set strDictionary = server.CreateObject("scripting.dictionary")

	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = "{keyicms:ProductList([\s\S]*?)}([\s\S]*?){/keyicms:ProductList}"

	set matches = regExpObj.Execute(Content)
  	for each match in matches
		labelStr = match.SubMatches(0)
		loopStr = match.SubMatches(1)
		
		set labelArr = ParseArr(labelStr)
		lNum = labelArr("Num")

		lSortID = labelArr("SortID")
		lSortPath = labelArr("SortPath")
		If lSortID = "" and lSortPath = "" Then
			lSortPath = "0,"
		ElseIf lSortID = "" and lSortPath <> "" Then
			lSortPath = lSortPath
		ElseIf lSortID <> "" Then
			sql = "select * from keyicms_ProductSort where ID="&lSortID
			If Conn.execute(sql).Eof Then
				ProductList = "SortID"&Keyicms_Lang_ErrMsg(0)
				Exit Function
			Else
				lSortPath = Conn.execute(sql)("SortPath")
			End If
		End If

		lCommendFlag = labelArr("CommendFlag")
		lNewFlag = labelArr("NewFlag")
		lOrder = labelArr("Order")
		lPageFlag = labelArr("PageFlag")
	
		If lPageFlag Then
			If IsNul(SortPath) Then
				dataSortPath = ""
			Else
				dataSortPath = " and Instr(SortPath,'"&SortPath&"')>0"
			End If
		Else
			If IsNul(lSortPath) Then
				dataSortPath = ""
			Else
				dataSortPath = " and Instr(SortPath,'"&lSortPath&"')>0"
			End If
		End If

		dim HideSort
		Set rs = server.CreateObject("adodb.recordset")
		rs.open "select * from  [keyicms_ProductSort] Where not ViewFlag" & dataSortPath, Conn, 1, 2
		while not rs.eof
			HideSort = "and not(Instr(SortPath,'"&rs("SortPath")&"')>0) " & HideSort
			rs.movenext
		wend
		rs.close : set rs = nothing

		If Ke01 = "PC" Then datawhere = " where ViewFlag"
		If Ke01 = "Mobile" Then datawhere = " where MobileFlag"
		
		datawhere = datawhere & dataSortPath
		If lCommendFlag = "True" Then
			datawhere = datawhere & " and CommendFlag"
		ElseIf lPicFlag = "False" Then
			datawhere = datawhere & " and not CommendFlag"
		End IF
		If lNewFlag = "True" Then
			datawhere = datawhere & " and NewFlag"
		ElseIf lNoticeFlag = "False" Then
			datawhere = datawhere & " and not NewFlag"
		End IF
	
		If lOrder="AddTime" Then
			taxis = " Order by "&lOrder&" Desc"
		ElseIf lOrder="Sequence" Then
			taxis = " Order by Sequence Asc, ID Desc"
		Else
			taxis = " Order by ID Desc"
		End If

		datafrom = "keyicms_Product"
		Pages = lNum

		Dim rs, sql
		Set rs = server.CreateObject("adodb.recordset")
		sql = "select * from ["& datafrom &"]" & datawhere & HideSort
		rs.Open sql, Conn, 1, 2
		idCount = rs.recordcount
		rs.close : set rs = nothing
			
		If idCount > 0 Then
			If(idCount Mod Pages = 0)Then
				Pagec = Int(idCount / Pages)
			Else
				Pagec = Int(idCount / Pages) + 1
			End If
			Set rs = server.CreateObject("adodb.recordset")
			sql = "select * from ["& datafrom &"] " & datawhere & HideSort & taxis
			rs.Open sql, Conn, 1, 1
			rs.Pagesize = Pages
			If Page < 1 Then Page = 1
			If Page > Pagec Then Page = Pagec
			If Pagec > 0 Then rs.absolutePage = Page
			For j = 1 To rs.Pagesize
				If rs.EOF Then Exit For
				If(j = 1)Then
					sqlid = rs("ID")
				Else
					sqlid = sqlid &","&rs("ID")
				End If
				rs.movenext
			Next
			rs.close : set rs = nothing
		End If
			
		Set regExpTwo = New Regexp
		regExpTwo.IgnoreCase = True
		regExpTwo.Global = True
		regExpTwo.Pattern = "{ProductList:([\s\S]*?)}"
		set matchesfield = regExpTwo.Execute(loopStr)
	
		If idCount > 0 And sqlid <> "" Then
			Set rs = server.CreateObject("adodb.recordset")
			sql = "select * from ["& datafrom &"] where ID in("& sqlid &") " & taxis
			rs.open sql, Conn, 1, 2
			loopstrTotal = ""
			For i = 1 to rs.recordcount
				If Ke01 = "PC" Then AutoLink = "/"&ProDir&"/"&rs("ClassSeo")&Separated&rs("ID")&"."&HTMLName
				If Ke01 = "Mobile" Then AutoLink = "/m/"&ProDir&"/"&rs("ClassSeo")&Separated&rs("ID")&"."&HTMLName
				nloopstr = loopStr
				for each matchfield in matchesfield
					FieldNameArr = regExpReplace(matchfield.SubMatches(0), "[\s]+", chr(32))
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
						Case "i"
							nloopstr = replace(nloopstr, matchfield.value, i)
						Case "Link"
							nloopstr = replace(nloopstr, matchfield.value, AutoLink)
						Case "SortName"
							SortName = Conn.execute("Select * from keyicms_ProductSort where ID="&rs("SortID"))("SortName")
							nloopstr = replace(nloopstr, matchfield.value, SortName)
						Case "SortID"
							nloopstr = replace(nloopstr, matchfield.value, rs("SortID"))
						Case "ID"
							nloopstr = replace(nloopstr, matchfield.value, rs("ID"))
						Case "ProductName"
							NameLen = ParseArr(FieldArr)("Len") 
							If cInt(NameLen) > 0 Then
								NameLen = cInt(NameLen)
								Title = StrLeft(rs("ProductName"), NameLen)
							Else
								Title = rs("ProductName")
							End If
							nloopstr = replace(nloopstr, matchfield.value, Title)
						Case "ProductNo"
							nloopstr = replace(nloopstr, matchfield.value, rs("ProductNo"))
						Case "Maker"
							nloopstr = replace(nloopstr, matchfield.value, rs("Maker"))
						Case "SmallPic"
							nloopstr = replace(nloopstr, matchfield.value, rs("SmallPic"))
						Case "BigPic"
							nloopstr = replace(nloopstr, matchfield.value, rs("BigPic"))
						Case "ClickNumber"
							nloopstr = replace(nloopstr, matchfield.value, rs("ClickNumber"))
						Case "AddTime"
							Format = ParseArr(FieldArr)("Format")
							nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("AddTime"), Format))
						Case "UpdateTime"
							Format = ParseArr(FieldArr)("Format")
							nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("UpdateTime"), Format))
						Case "Content"
							ContLen = ParseArr(FieldArr)("Len")
							If cInt(ContLen) > 0 Then
								ContLen = cInt(ContLen)
								Cont = StrLeft(RemoveHTML(rs("Content")), ContLen)
							Else
								Cont = RemoveHTML(rs("Content"))
							End If
							nloopstr = replace(nloopstr, matchfield.value, Cont)
					end select
				next
				loopstrTotal = loopstrTotal & nloopstr
				rs.movenext
			Next
			rs.close : set rs = nothing
		Else
			loopstrTotal = Keyicms_Lang_ErrMsg(1)
		End If
		set matchesfield = nothing
		Content = replace(Content, match.value, loopstrTotal)
		strDictionary.removeAll
		If lPageFlag = "True" Then
			Content = ShowPage(Content, idCount, Page, Pagec, Pages)'内容，记录总数，当前页数，共页数，每页记录数
		End If
	next
	set matches = nothing
	ProductList = Content
End Function

Function CaseList(Content, Page, Ke01, SortPath)
	If Page = "" Then Page = 1
	Set regExpObj = New Regexp
	set strDictionary = server.CreateObject("scripting.dictionary")

	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = "{keyicms:CaseList([\s\S]*?)}([\s\S]*?){/keyicms:CaseList}"

	set matches = regExpObj.Execute(Content)
  	for each match in matches
		labelStr = match.SubMatches(0)
		loopStr = match.SubMatches(1)
		
		set labelArr = ParseArr(labelStr)
		lNum = labelArr("Num")

		lSortID = labelArr("SortID")
		lSortPath = labelArr("SortPath")
		If lSortID = "" and lSortPath = "" Then
			lSortPath = "0,"
		ElseIf lSortID = "" and lSortPath <> "" Then
			lSortPath = lSortPath
		ElseIf lSortID <> "" Then
			sql = "select * from keyicms_CaseSort where ID="&lSortID
			If Conn.execute(sql).Eof Then
				CaseList = "SortID"&Keyicms_Lang_ErrMsg(0)
				Exit Function
			Else
				lSortPath = Conn.execute(sql)("SortPath")
			End If
		End If

		lCommendFlag = labelArr("CommendFlag")
		lOrder = labelArr("Order")
		lPageFlag = labelArr("PageFlag")

		If lPageFlag Then
			If IsNul(SortPath) Then
				dataSortPath = ""
			Else
				dataSortPath = " and Instr(SortPath,'"&SortPath&"')>0"
			End If
		Else
			If IsNul(lSortPath) Then
				dataSortPath = ""
			Else
				dataSortPath = " and Instr(SortPath,'"&lSortPath&"')>0"
			End If
		End If

		dim HideSort
		Set rs = server.CreateObject("adodb.recordset")
		rs.open "select * from  [keyicms_CaseSort] Where not ViewFlag" & dataSortPath, Conn, 1, 2
		while not rs.eof
			HideSort = "and not(Instr(SortPath,'"&rs("SortPath")&"')>0) " & HideSort
			rs.movenext
		wend
		rs.close : set rs = nothing
		
		If Ke01 = "PC" Then datawhere = " where ViewFlag"
		If Ke01 = "Mobile" Then datawhere = " where MobileFlag"
		
		datawhere = datawhere & dataSortPath
		If lCommendFlag = "True" Then
			datawhere = datawhere & " and CommendFlag"
		ElseIf lPicFlag = "False" Then
			datawhere = datawhere & " and not CommendFlag"
		End IF
	
		If lOrder="AddTime" Then
			taxis = " Order by "&lOrder&" Desc"
		ElseIf lOrder="Sequence" Then
			taxis = " Order by Sequence Asc, ID Desc"
		Else
			taxis = " Order by ID Desc"
		End If

		datafrom = "keyicms_Case"
		Pages = lNum
		
		Dim rs, sql
		Set rs = server.CreateObject("adodb.recordset")
		sql = "select * from ["& datafrom &"]" & datawhere & HideSort
		rs.Open sql, Conn, 1, 2
		idCount = rs.recordcount
		rs.close : set rs = nothing
	
		If idCount > 0 Then
			If(idCount Mod Pages = 0)Then
				Pagec = Int(idCount / Pages)
			Else
				Pagec = Int(idCount / Pages) + 1
			End If
			Set rs = server.CreateObject("adodb.recordset")
			sql = "select * from ["& datafrom &"] " & datawhere & HideSort & taxis
			rs.Open sql, Conn, 1, 1
			rs.Pagesize = Pages
			If Page < 1 Then Page = 1
			If Page > Pagec Then Page = Pagec
			If Pagec > 0 Then rs.absolutePage = Page
			For j = 1 To rs.Pagesize
				If rs.EOF Then Exit For
				If(j = 1)Then
					sqlid = rs("ID")
				Else
					sqlid = sqlid &","&rs("ID")
				End If
				rs.movenext
			Next
			rs.close : set rs = nothing
		End If
			
		Set regExpTwo = New Regexp
		regExpTwo.IgnoreCase = True
		regExpTwo.Global = True
		regExpTwo.Pattern = "{CaseList:([\s\S]*?)}"
		set matchesfield = regExpTwo.Execute(loopStr)
	
		If idCount > 0 And sqlid <> "" Then
			Set rs = server.CreateObject("adodb.recordset")
			sql = "select * from ["& datafrom &"] where ID in("& sqlid &") " & taxis
			rs.open sql, Conn, 1, 2
			loopstrTotal = ""
			For i = 1 to rs.recordcount
				If Ke01 = "PC" Then AutoLink = "/"&CaseDir&"/"&rs("ClassSeo")&Separated&rs("ID")&"."&HTMLName
				If Ke01 = "Mobile" Then AutoLink = "/m/"&CaseDir&"/"&rs("ClassSeo")&Separated&rs("ID")&"."&HTMLName
				nloopstr = loopStr
				for each matchfield in matchesfield
					FieldNameArr = regExpReplace(matchfield.SubMatches(0), "[\s]+", chr(32))
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
						Case "i"
							nloopstr = replace(nloopstr, matchfield.value, i)
						Case "Link"
							nloopstr = replace(nloopstr, matchfield.value, AutoLink)
						Case "SortName"
							SortName = Conn.execute("Select * from keyicms_CaseSort where ID="&rs("SortID"))("SortName")
							nloopstr = replace(nloopstr, matchfield.value, SortName)
						Case "SortID"
							nloopstr = replace(nloopstr, matchfield.value, rs("SortID"))
						Case "ID"
							nloopstr = replace(nloopstr, matchfield.value, rs("ID"))
						Case "CaseName"
							NameLen = ParseArr(FieldArr)("Len") 
							If cInt(NameLen) > 0 Then
								NameLen = cInt(NameLen)
								Title = StrLeft(rs("CaseName"), NameLen)
							Else
								Title = rs("CaseName")
							End If
							nloopstr = replace(nloopstr, matchfield.value, title)
						Case "SmallPic"
							nloopstr = replace(nloopstr, matchfield.value, rs("SmallPic"))
						Case "BigPic"
							nloopstr = replace(nloopstr, matchfield.value, rs("BigPic"))
						Case "AddTime"
							Format = ParseArr(FieldArr)("Format")
							nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("AddTime"), Format))
						Case "UpdateTime"
							Format = ParseArr(FieldArr)("Format")
							nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("UpdateTime"), Format))
						Case "ClickNumber"
							nloopstr = replace(nloopstr, matchfield.value, rs("ClickNumber"))
						Case "Content"
							ContLen = ParseArr(FieldArr)("Len")
							If cInt(ContLen) > 0 Then
								ContLen = cInt(ContLen)
								Cont = StrLeft(RemoveHTML(rs("Content")), ContLen)
							Else
								Cont = RemoveHTML(rs("Content"))
							End If
							nloopstr = replace(nloopstr, matchfield.value, Cont)
					end select
				next
				loopstrTotal = loopstrTotal & nloopstr
				rs.movenext
			Next
			rs.close : set rs = nothing
		Else
			loopstrTotal = Keyicms_Lang_ErrMsg(1)
		End If
		set matchesfield = nothing
		Content = replace(Content, match.value, loopstrTotal)
		strDictionary.removeAll
		If lPageFlag = "True" Then
			Content = ShowPage(Content, idCount, Page, Pagec, Pages)'内容，记录总数，当前页数，共页数，每页记录数
		End If
	next
	set matches = nothing
	CaseList = Content
End Function

Function DownList(Content, Page, Ke01, SortPath)
	If Page = "" Then Page = 1
	Set regExpObj = New Regexp
	set strDictionary = server.CreateObject("scripting.dictionary")

	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = "{keyicms:DownList([\s\S]*?)}([\s\S]*?){/keyicms:DownList}"

	set matches = regExpObj.Execute(Content)
  	for each match in matches
		labelStr = match.SubMatches(0)
		loopStr = match.SubMatches(1)
		
		set labelArr = ParseArr(labelStr)
		lNum = labelArr("Num")

		lSortID = labelArr("SortID")
		lSortPath = labelArr("SortPath")
		If lSortID = "" and lSortPath = "" Then
			lSortPath = "0,"
		ElseIf lSortID = "" and lSortPath <> "" Then
			lSortPath = lSortPath
		ElseIf lSortID <> "" Then
			sql = "select * from keyicms_DownSort where ID="&lSortID
			If Conn.execute(sql).Eof Then
				DownList = "SortID"&Keyicms_Lang_ErrMsg(0)
				Exit Function
			Else
				lSortPath = Conn.execute(sql)("SortPath")
			End If
		End If

		lOrder = labelArr("Order")
		lPageFlag = labelArr("PageFlag")
	
		If lPageFlag Then
			If IsNul(SortPath) Then
				dataSortPath = ""
			Else
				dataSortPath = " and Instr(SortPath,'"&SortPath&"')>0"
			End If
		Else
			If IsNul(lSortPath) Then
				dataSortPath = ""
			Else
				dataSortPath = " and Instr(SortPath,'"&lSortPath&"')>0"
			End If
		End If

		dim HideSort
		Set rs = server.CreateObject("adodb.recordset")
		rs.open "select * from  [keyicms_DownSort] Where not ViewFlag" & dataSortPath, Conn, 1, 2
		while not rs.eof
			HideSort = "and not(Instr(SortPath,'"&rs("SortPath")&"')>0) " & HideSort
			rs.movenext
		wend
		rs.close : set rs = nothing

		If Ke01 = "PC" Then datawhere = " where ViewFlag"
		If Ke01 = "Mobile" Then datawhere = " where MobileFlag"
		
		datawhere = datawhere & dataSortPath
		If lOrder="AddTime" Then
			taxis = " Order by "&lOrder&" Desc"
		Else
			taxis = " Order by ID Desc"
		End If

		datafrom = "keyicms_Download"
		Pages = lNum
		
		Dim rs, sql
		Set rs = server.CreateObject("adodb.recordset")
		sql = "select * from ["& datafrom &"]" & datawhere & HideSort
		rs.Open sql, Conn, 1, 2
		idCount = rs.recordcount
		rs.close : set rs = nothing
	
		If idCount > 0 Then
			If(idCount Mod Pages = 0)Then
				Pagec = Int(idCount / Pages)
			Else
				Pagec = Int(idCount / Pages) + 1
			End If
			Set rs = server.CreateObject("adodb.recordset")
			sql = "select * from ["& datafrom &"] " & datawhere & HideSort & taxis
			rs.Open sql, Conn, 1, 1
			rs.Pagesize = Pages
			If Page < 1 Then Page = 1
			If Page > Pagec Then Page = Pagec
			If Pagec > 0 Then rs.absolutePage = Page
			For j = 1 To rs.Pagesize
				If rs.EOF Then Exit For
				If(j = 1)Then
					sqlid = rs("ID")
				Else
					sqlid = sqlid &","&rs("ID")
				End If
				rs.movenext
			Next
			rs.close : set rs = nothing
		End If
			
		Set regExpTwo = New Regexp
		regExpTwo.IgnoreCase = True
		regExpTwo.Global = True
		regExpTwo.Pattern = "{DownList:([\s\S]*?)}"
		set matchesfield = regExpTwo.Execute(loopStr)
	
		If idCount > 0 And sqlid <> "" Then
			Set rs = server.CreateObject("adodb.recordset")
			sql = "select * from ["& datafrom &"] where ID in("& sqlid &") " & taxis
			rs.open sql, Conn, 1, 2
			loopstrTotal = ""
			For i = 1 to rs.recordcount
				If Ke01 = "PC" Then AutoLink = "/"&DownDir&"/"&rs("ClassSeo")&Separated&rs("ID")&"."&HTMLName
				If Ke01 = "Mobile" Then AutoLink = "/m/"&DownDir&"/"&rs("ClassSeo")&Separated&rs("ID")&"."&HTMLName
				nloopstr = loopStr
				for each matchfield in matchesfield
					FieldNameArr = regExpReplace(matchfield.SubMatches(0), "[\s]+", chr(32))
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
						Case "i"
							nloopstr = replace(nloopstr, matchfield.value, i)
						Case "Link"
							nloopstr = replace(nloopstr, matchfield.value, AutoLink)
						Case "SortName"
							SortName = Conn.execute("Select * from keyicms_DownSort where ID="&rs("SortID"))("SortName")
							nloopstr = replace(nloopstr, matchfield.value, SortName)
						Case "SortID"
							nloopstr = replace(nloopstr, matchfield.value, rs("SortID"))
						Case "ID"
							nloopstr = replace(nloopstr, matchfield.value, rs("ID"))
						Case "DownName"
							NameLen = ParseArr(FieldArr)("Len") 
							If cInt(NameLen) > 0 Then
								NameLen = cInt(NameLen)
								Title = StrLeft(rs("DownName"), NameLen)
							Else
								Title = rs("DownName")
							End If
							nloopstr = replace(nloopstr, matchfield.value, Title)
						Case "FileSize"
							nloopstr = replace(nloopstr, matchfield.value, rs("FileSize"))
						Case "FileUrl"
							nloopstr = replace(nloopstr, matchfield.value, rs("FileUrl"))
						Case "AddTime"
							Format = ParseArr(FieldArr)("Format")
							nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("AddTime"), Format))
						Case "UpdateTime"
							Format = ParseArr(FieldArr)("Format")
							nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("UpdateTime"), Format))
						Case "ClickNumber"
							nloopstr = replace(nloopstr, matchfield.value, rs("ClickNumber"))
						Case "Content"
							ContLen = ParseArr(FieldArr)("Len")
							If cInt(ContLen) > 0 Then
								ContLen = cInt(ContLen)
								Cont = StrLeft(RemoveHTML(rs("Content")), ContLen)
							Else
								Cont = RemoveHTML(rs("Content"))
							End If
							nloopstr = replace(nloopstr, matchfield.value, Cont)
					end select
				next
				loopstrTotal = loopstrTotal & nloopstr
				rs.movenext
			Next
			rs.close : set rs = nothing
		Else
			loopstrTotal = Keyicms_Lang_ErrMsg(1)
		End If
		set matchesfield = nothing
		Content = replace(Content, match.value, loopstrTotal)
		strDictionary.removeAll
		If lPageFlag = "True" Then
			Content = ShowPage(Content, idCount, Page, Pagec, Pages)
		End If
	next
	set matches = nothing
	DownList = Content
End Function

Function JobList(Content, Page, Ke01)
	If Page = "" Then Page = 1
	Set regExpObj = New Regexp
	set strDictionary = server.CreateObject("scripting.dictionary")

	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = "{keyicms:JobList([\s\S]*?)}([\s\S]*?){/keyicms:JobList}"

	set matches = regExpObj.Execute(Content)
  	for each match in matches
		labelStr = match.SubMatches(0)
		loopStr = match.SubMatches(1)
		
		set labelArr = ParseArr(labelStr)
		lNum = labelArr("Num")
		lOrder = labelArr("Order")
		lPageFlag = labelArr("PageFlag")
	
		If Ke01 = "PC" Then datawhere = " where ViewFlag"
		If Ke01 = "Mobile" Then datawhere = " where MobileFlag"

		If lOrder="AddTime" Then
			taxis = " Order by "&lOrder&" Desc"
		Else
			taxis = " Order by ID Desc"
		End If

		datafrom = "keyicms_Job"
		Pages = lNum
		
		Dim rs, sql
		Set rs = server.CreateObject("adodb.recordset")
		sql = "select * from ["& datafrom &"]" & datawhere
		rs.Open sql, Conn, 1, 2
		idCount = rs.recordcount
		rs.close : set rs = nothing
	
		If idCount > 0 Then
			If(idCount Mod Pages = 0)Then
				Pagec = Int(idCount / Pages)
			Else
				Pagec = Int(idCount / Pages) + 1
			End If
			Set rs = server.CreateObject("adodb.recordset")
			sql = "select * from ["& datafrom &"] " & datawhere
			rs.Open sql, Conn, 1, 1
			rs.Pagesize = Pages
			If Page < 1 Then Page = 1
			If Page > Pagec Then Page = Pagec
			If Pagec > 0 Then rs.absolutePage = Page
			For j = 1 To rs.Pagesize
				If rs.EOF Then Exit For
				If(j = 1)Then
					sqlid = rs("ID")
				Else
					sqlid = sqlid &","&rs("ID")
				End If
				rs.movenext
			Next
			rs.close : set rs = nothing
		End If
			
		Set regExpTwo = New Regexp
		regExpTwo.IgnoreCase = True
		regExpTwo.Global = True
		regExpTwo.Pattern = "{JobList:([\s\S]*?)}"
		set matchesfield = regExpTwo.Execute(loopStr)
	
		If idCount > 0 And sqlid <> "" Then
			Set rs = server.CreateObject("adodb.recordset")
			sql = "select * from ["& datafrom &"] where ID in("& sqlid &") " & taxis
			rs.open sql, Conn, 1, 2
			loopstrTotal = ""
			For i = 1 to rs.recordcount
				If Ke01 = "PC" Then AutoLink = "/"&JobDir&"/"&rs("ClassSeo")&Separated&rs("ID")&"."&HTMLName
				If Ke01 = "Mobile" Then AutoLink = "/m/"&JobDir&"/"&rs("ClassSeo")&Separated&rs("ID")&"."&HTMLName
				nloopstr = loopStr
				for each matchfield in matchesfield
					FieldNameArr = regExpReplace(matchfield.SubMatches(0), "[\s]+", chr(32))
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
						Case "i"
							nloopstr = replace(nloopstr, matchfield.value, i)
						Case "Link"
							nloopstr = replace(nloopstr, matchfield.value, AutoLink)
						Case "ID"
							nloopstr = replace(nloopstr, matchfield.value, rs("ID"))
						Case "JobName"
							NameLen = ParseArr(FieldArr)("Len") 
							If cInt(NameLen) > 0 Then
								NameLen = cInt(NameLen)
								Title = StrLeft(rs("JobName"), NameLen)
							Else
								Title = rs("JobName")
							End If
							nloopstr = replace(nloopstr, matchfield.value, Title)
						Case "JobAddress"
							nloopstr = replace(nloopstr, matchfield.value, rs("JobAddress"))
						Case "JobNumber"
							nloopstr = replace(nloopstr, matchfield.value, rs("JobNumber"))
						Case "Emolument"
							nloopstr = replace(nloopstr, matchfield.value, rs("Emolument"))
						Case "StartDate"
							Format = ParseArr(FieldArr)("Format")
							nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("StartDate"), Format))
						Case "EndDate"
							Format = ParseArr(FieldArr)("Format")
							nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("EndDate"), Format))
						Case "AddTime"
							Format = ParseArr(FieldArr)("Format")
							nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("AddTime"), Format))
						Case "UpdateTime"
							Format = ParseArr(FieldArr)("Format")
							nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("UpdateTime"), Format))
						Case "Requirement"
							ContLen = ParseArr(FieldArr)("Len") 
							If cInt(ContLen) > 0 Then
								ContLen = cInt(ContLen)
								Cont = StrLeft(RemoveHTML(rs("Requirement")), ContLen)
							Else
								Cont = RemoveHTML(rs("Requirement"))
							End If
							nloopstr = replace(nloopstr, matchfield.value, Cont)
						Case "ClickNumber"
							nloopstr = replace(nloopstr, matchfield.value, rs("ClickNumber"))
					end select
				next
				loopstrTotal = loopstrTotal & nloopstr
				rs.movenext
			Next
			rs.close : set rs = nothing
		Else
			loopstrTotal = Keyicms_Lang_ErrMsg(1)
		End If
		set matchesfield = nothing
		Content = replace(Content, match.value, loopstrTotal)
		strDictionary.removeAll
		If lPageFlag = "True" Then
			Content = ShowPage(Content, idCount, Page, Pagec, Pages)
		End If
	next
	set matches = nothing
	JobList = Content
End Function

Function Ads(labelStr, loopStr)
	set labelArr = ParseArr(labelStr)
	lSortID = labelArr("SortID")
	lSortPath = labelArr("SortPath")
	If lSortID = "" and lSortPath = "" Then
		lSortPath = "0,"
	ElseIf lSortID = "" and lSortPath <> "" Then
		lSortPath = lSortPath
	ElseIf lSortID <> "" Then
		sql = "select * from keyicms_AdsSort where ID="&lSortID
		If Conn.execute(sql).Eof Then
			Ads = "SortID"&Keyicms_Lang_ErrMsg(0)
			Exit Function
		Else
			lSortPath = Conn.execute(sql)("SortPath")
		End If
	End If
	lID = labelArr("ID")

	Set regExpTwo = New Regexp
	regExpTwo.IgnoreCase = True
	regExpTwo.Global = True
	regExpTwo.Pattern = "{Ads:([\s\S]*?)}"
	set matchesfield = regExpTwo.Execute(loopStr)
	
	Set rs = server.CreateObject("adodb.recordset")
	If lID<>"" Then
		sql = "select * from keyicms_Ads where ID="&lID
	Else
		sql = "select * from keyicms_Ads where Instr(SortPath, '"&lSortPath&"')>0 Order by Sequence Asc, ID Desc"
	End If

	rs.open sql, Conn, 1, 2
	loopstrTotal = ""
	If rs.eof Then
		loopstrTotal = Keyicms_Lang_ErrMsg(1)
	Else
		do until rs.eof
			nloopstr = loopStr
			for each matchfield in matchesfield
				FieldNameArr = regExpReplace(matchfield.SubMatches(0), "[\s]+", chr(32))
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
					Case "ID"
						nloopstr = replace(nloopstr, matchfield.value, rs("ID"))
					Case "SortName"
						SortName = Conn.execute("Select * from keyicms_AdsSort where ID="&rs("SortID"))("SortName")
						nloopstr = replace(nloopstr, matchfield.value, SortName)
					Case "AdsUrl"
						nloopstr = replace(nloopstr, matchfield.value, rs("AdsUrl"))
					Case "AdsName"
						nloopstr = replace(nloopstr, matchfield.value, rs("AdsName"))
					Case "AdsPic"
						nloopstr = replace(nloopstr, matchfield.value, rs("AdsPic"))
					Case "Content"
						ContLen = ParseArr(FieldArr)("Len") 
						If cInt(ContLen) > 0 Then
							ContLen = cInt(ContLen)
							Cont = StrLeft(RemoveHTML(rs("Content")), ContLen)
						Else
							Cont = RemoveHTML(rs("Content"))
						End If
						nloopstr = replace(nloopstr, matchfield.value, Cont)
				end select
			next
			loopstrTotal = loopstrTotal & nloopstr
			rs.movenext
		loop
		rs.close : set rs = nothing
		set matchesfield = nothing
	End If
	Ads = loopstrTotal
End Function

Function FriendLink(labelStr, loopStr)
	set labelArr = ParseArr(labelStr)
	lLinkType = labelArr("LinkType")
	
	Set regExpTwo = New Regexp
	regExpTwo.IgnoreCase = True
	regExpTwo.Global = True
	regExpTwo.Pattern = "{FriendLink:([\s\S]*?)}"
	set matchesfield = regExpTwo.Execute(loopStr)
		
	Set rs = server.CreateObject("adodb.recordset")
	sql = "select * from keyicms_FriendLink Where LinkType = "&lLinkType&" Order by ID Desc"
	rs.open sql, Conn, 1, 2
	loopstrTotal = ""
	If rs.eof Then
		loopstrTotal = Keyicms_Lang_ErrMsg(1)
	Else
		do until rs.eof
			nloopstr = loopStr
			for each matchfield in matchesfield
				FieldNameArr = regExpReplace(matchfield.SubMatches(0), "[\s]+", chr(32))
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
					Case "ID"
						nloopstr = replace(nloopstr, matchfield.value, rs("ID"))
					Case "LinkUrl"
						nloopstr = replace(nloopstr, matchfield.value, rs("LinkUrl"))
					Case "LinkFace"
						nloopstr = replace(nloopstr, matchfield.value, rs("LinkFace"))
					Case "LinkName"
						nloopstr = replace(nloopstr, matchfield.value, rs("LinkName"))
					Case "Remark"
						nloopstr = replace(nloopstr, matchfield.value, rs("Remark"))
					Case "AddTime"
						Format = ParseArr(FieldArr)("Format")
						nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("AddTime"), Format))
				end select
			next
			loopstrTotal = loopstrTotal & nloopstr
			rs.movenext
		loop
		rs.close : set rs = nothing
		set matchesfield = nothing
	End If
	FriendLink = loopstrTotal
End Function

Function Slide(loopStr)
	Set regExpTwo = New Regexp
	regExpTwo.IgnoreCase = True
	regExpTwo.Global = True
	regExpTwo.Pattern = "{Slide:([\s\S]*?)}"
	set matchesfield = regExpTwo.Execute(loopStr)

	Set rs = server.CreateObject("adodb.recordset")
	sql = "select * from keyicms_Slide Order by State Asc, ID Desc"
	rs.open sql, Conn, 1, 2
	loopstrTotal = ""
	do until rs.eof
		nloopstr = loopStr
		for each matchfield in matchesfield
			FieldNameArr = regExpReplace(matchfield.SubMatches(0), "[\s]+", chr(32))
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
				Case "ID"
					nloopstr = replace(nloopstr, matchfield.value, rs("ID"))
				Case "Url"
					nloopstr = replace(nloopstr, matchfield.value, rs("Url"))
				Case "Pic"
					nloopstr = replace(nloopstr, matchfield.value, rs("Pic"))
				Case "SlideName"
					nloopstr = replace(nloopstr, matchfield.value, rs("SlideName"))
			end select
		next
		loopstrTotal = loopstrTotal & nloopstr
		rs.movenext
	loop
	rs.close : set rs = nothing
	set matchesfield = nothing
	
	Slide = loopstrTotal
End Function

Function Contact(labelStr, loopStr)
	set labelArr = ParseArr(labelStr)
	lStatus = labelArr("Status")
	If not isNum(lStatus) Then Contact = "" : exit Function

	Set regExpTwo = New Regexp
	regExpTwo.IgnoreCase = True
	regExpTwo.Global = True
	regExpTwo.Pattern = "{Contact:([\s\S]*?)}"
	set matchesfield = regExpTwo.Execute(loopStr)

	Set rs = server.CreateObject("adodb.recordset")
	If lStatus = 1 Then
		sql = "select * from keyicms_Contact where SortID=1 and ID<>1 Order by Sequence Asc, ID Desc"
	ElseIf lStatus = 2 Then
		sql = "select * from keyicms_Contact where SortID=2"
	End If
	rs.open sql, Conn, 1, 2
	loopstrTotal = ""
	do until rs.eof
		nloopstr = loopStr
		for each matchfield in matchesfield
			FieldNameArr = regExpReplace(matchfield.SubMatches(0), "[\s]+", chr(32))
			FieldNameArr = trim(FieldNameArr)
			m = instr(FieldNameArr, chr(32))
			if  m > 0 then 
				FieldName = left(FieldNameArr, m - 1)
				FieldArr =	right(FieldNameArr, len(FieldNameArr) - m)
			else
				FieldArr =	""
				FieldName = FieldNameArr
			end if
			Select Case FieldName
				Case "ID"
					nloopstr = replace(nloopstr, matchfield.value, rs("ID"))
				Case "Code"
					nloopstr = replace(nloopstr, matchfield.value, rs("Code"))
				Case "Qq"
					nloopstr = replace(nloopstr, matchfield.value, changeNull(rs("Qq")))
				Case "User"
					nloopstr = replace(nloopstr, matchfield.value, rs("User"))
				Case "Phone"
					nloopstr = replace(nloopstr, matchfield.value, changeNull(rs("Phone")))
				Case "WeChat"
					nloopstr = replace(nloopstr, matchfield.value, changeNull(rs("WeChat")))
			end select
		next
		loopstrTotal = loopstrTotal & nloopstr
		rs.movenext
	loop
	rs.close : set rs = nothing
	set matchesfield = nothing
	
	Contact = loopstrTotal
End Function

Function Search(Content, Page, Ke01, QUERY_STRING)
	If Page = "" Then Page = 1
	Set regExpObj = New Regexp
	set strDictionary = server.CreateObject("scripting.dictionary")

	regExpObj.IgnoreCase = True
	regExpObj.Global = True
	regExpObj.Pattern = "{keyicms:Search([\s\S]*?)}([\s\S]*?){/keyicms:Search}"
	set matches = regExpObj.Execute(Content)

  	for each match in matches
		labelStr = match.SubMatches(0)
		loopStr = match.SubMatches(1)

		set labelArr = ParseArr(labelStr)
		lNum = labelArr("Num")
		Pages = lNum

		If Ke01 = "PC" Then Flag = "ViewFlag"
		If Ke01 = "Mobile" Then Flag = "MobileFlag"

		Dim rs, sql
		set rs = server.CreateObject("adodb.recordset")
		rs.open "select * from keyicms_System where ViewFlag Order by DataForm Asc", Conn, 1, 2
		For i = 1 to rs.recordcount
If rs("DataForm") = 1 Then sql = sql & "select '"&rs("DataForm")&"' as DataForm, ID, AddTime, ClassSeo, AboutName as Title   from keyicms_About    where "&Flag&" and AboutName   like '%"&key&"%' Order by ID Desc " & " union "
If rs("DataForm") = 2 Then sql = sql & "select '"&rs("DataForm")&"' as DataForm, ID, AddTime, ClassSeo, NewsName as Title    from keyicms_News     where "&Flag&" and NewsName    like '%"&key&"%' Order by ID Desc " & " union "
If rs("DataForm") = 3 Then sql = sql & "select '"&rs("DataForm")&"' as DataForm, ID, AddTime, ClassSeo, ProductName as Title from keyicms_Product  where "&Flag&" and ProductName like '%"&key&"%' Order by ID Desc " & " union "
If rs("DataForm") = 4 Then sql = sql & "select '"&rs("DataForm")&"' as DataForm, ID, AddTime, ClassSeo, CaseName as Title    from keyicms_Case     where "&Flag&" and CaseName    like '%"&key&"%' Order by ID Desc " & " union "
If rs("DataForm") = 5 Then sql = sql & "select '"&rs("DataForm")&"' as DataForm, ID, AddTime, ClassSeo, DownName as Title    from keyicms_Download where "&Flag&" and DownName    like '%"&key&"%' Order by ID Desc " & " union "
If rs("DataForm") = 6 Then sql = sql & "select '"&rs("DataForm")&"' as DataForm, ID, AddTime, ClassSeo, JobName as Title     from keyicms_Job      where "&Flag&" and JobName     like '%"&key&"%' Order by ID Desc " & " union "
rs.movenext
		Next
		rs.close : set rs = nothing
		
		SqlArray = split(sql, "union")
		If Ubound(SqlArray) < 0 Then
			loopstrTotal = "未搜索到相关信息"
		Else
			For i = 0 to Ubound(SqlArray) - 1
				If i = 0 Then
					sql = sql & SqlArray(i)
				Else
					sql = sql & " union " & SqlArray(i)
				End If
			Next

			Set rs = server.CreateObject("adodb.recordset")
			sql = "Select * from (" & sql & ") Order by AddTime Desc"
			rs.Open sql, Conn, 1, 2
			idCount = rs.recordcount
			rs.close : set rs = nothing
			If idCount > 0 Then
				If(idCount Mod Pages = 0)Then
					Pagec = Int(idCount / Pages)
				Else
					Pagec = Int(idCount / Pages) + 1
				End If
				Set rs = server.CreateObject("adodb.recordset")
				rs.Open sql, Conn, 1, 1
				rs.Pagesize = Pages
				If Page < 1 Then Page = 1
				If Page > Pagec Then Page = Pagec
				If Pagec > 0 Then rs.absolutePage = Page
				Set regExpTwo = New Regexp
				regExpTwo.IgnoreCase = True
				regExpTwo.Global = True
				regExpTwo.Pattern = "{Search:([\s\S]*?)}"
				set matchesfield = regExpTwo.Execute(loopStr)
	
				loopstrTotal = ""
				For j = 1 To rs.Pagesize
					If rs.eof Then Exit For
					If Ke01 = "PC" Then
						mPath = ""
					Else
						mPath = "/m"
					End If
	
					nloopstr = loopStr
					for each matchfield in matchesfield
						FieldNameArr = regExpReplace(matchfield.SubMatches(0), "[\s]+", chr(32))
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
								SystemDir = Conn.execute("Select * from keyicms_System where DataForm="&rs("DataForm"))("SystemDir")
								AutoLink = mPath&"/"&SystemDir&"/"&rs("ClassSeo")&Separated&rs("ID")&"."&HTMLName
								nloopstr = replace(nloopstr, matchfield.value, AutoLink)
							Case "Title"
								NameLen = ParseArr(FieldArr)("Len") 
								If cInt(NameLen) > 0 Then
									NameLen = cInt(NameLen)
									Title = StrLeft(rs("Title"), NameLen)
								Else
									Title = rs("Title")
								End If
								nloopstr = replace(nloopstr, matchfield.value, Title)
							Case "AddTime"
								Format = ParseArr(FieldArr)("Format")
								nloopstr = replace(nloopstr, matchfield.value, FormatDate(rs("AddTime"), Format))
							Case "Column"
								DataForm = Conn.execute("Select * from keyicms_System where DataForm="&rs("DataForm"))("SeoTitle")
								nloopstr = replace(nloopstr, matchfield.value, DataForm)
						end select
					next
					loopstrTotal = loopstrTotal & nloopstr
					rs.movenext
				Next
				rs.close : set rs = nothing
			Else
				loopstrTotal = Keyicms_Lang_ErrMsg(4)
			End If
		End If
		set matchesfield = nothing
		Content = replace(Content, match.value, loopstrTotal)
		strDictionary.removeAll
	next
	set matches = nothing
	Content = ShowPageAsp(Content, idCount, Page, Pagec, Pages)
	Search = Content
End Function
%>