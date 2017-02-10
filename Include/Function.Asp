<%
Function clscache()
    Dim cacheobj
    application.Lock
    For Each cacheobj In application.contents
        If CStr(Left(cacheobj, Len(Cacheflag))) = CStr(Cacheflag) Then application.contents.Remove (cacheobj)
    Next
    application.unlock
End Function

Function setcache(cachename, cachevalue)
    Dim cachedata
    cachename = LCase(filterstr(cachename))
    cachedata = application(Cacheflag & cachename)
    If IsArray(cachedata) Then
        cachedata(0) = Cachevalue
        cachedata(1) = Now()
    Else
        ReDim cachedata(2)
        cachedata(0) = Cachevalue
        cachedata(1) = Now()
    End If
    application.Lock
    application(Cacheflag & cachename) = cachedata
    application.unlock
End Function

Function getcache(cachename)
    Dim cachedata
    cachename = LCase(filterstr(cachename))
    cachedata = application(Cacheflag & cachename)
    If IsArray(cachedata) Then getcache = cachedata(0) Else getcache = ""
End Function

Function chkcache(cachename)
    Dim cachedata
    chkcache = False
    cachename = LCase(filterstr(cachename))
    cachedata = application(Cacheflag & cachename)
    If Not IsArray(cachedata) Then Exit Function
    If Not IsDate(cachedata(1)) Then Exit Function
    If DateDiff("s", CDate(cachedata(1)), Now()) < 60 * Cachetime Then chkcache = True
End Function

Function filterstr(Str)
    filterstr = LCase(Str)
    filterstr = Replace(filterstr, " ", "")
    filterstr = Replace(filterstr, "'", "")
    filterstr = Replace(filterstr, """", "")
    filterstr = Replace(filterstr, "=", "")
    filterstr = Replace(filterstr, "*", "")
End Function

Function ReplaceBadChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceBadChar = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "+,',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ",--"
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    ReplaceBadChar = tempChar
End Function

Function ReplaceConstChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceConstChar = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "+,',%,^,&,?,(,),<,>,[,],{,},;,:," & Chr(34) & "," & Chr(0) & ",--"
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    ReplaceConstChar = tempChar
End Function

Function StrLen(Str)
    If Str = "" Or IsNull(Str) Then
        StrLen = 0
        Exit Function
    Else
        Dim regex
        Set regex = New regexp
        regEx.Pattern = "[^\x00-\xff]"
        regex.Global = True
        Str = regEx.Replace(Str, "^^")
        Set regex = Nothing
        StrLen = Len(Str)
    End If
End Function

Function ReStrReplace(Str)
    If Str = "" Or IsNull(Str) Then
        ReStrReplace = ""
        Exit Function
    Else
        ReStrReplace = Replace(Str, "&nbsp;", " ")
        ReStrReplace = Replace(ReStrReplace, "<br />", Chr(13))
        ReStrReplace = Replace(ReStrReplace, "&lt;br&gt;", Chr(13))
        ReStrReplace = Replace(ReStrReplace, "&lt;", "<")
        ReStrReplace = Replace(ReStrReplace, "&gt;", ">")
    End If
End Function

Function HtmlStrReplace(Str)
    If Str = "" Or IsNull(Str) Then
        HtmlStrReplace = ""
        Exit Function
    Else
        HtmlStrReplace = Replace(Str, "&lt;br&gt;", "<br />")
    End If
End Function

Function GetUrl()
    GetUrl = "http://"&Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL")
    If Request.ServerVariables("QUERY_STRING")<>"" Then GetURL = GetUrl&"?"& Request.ServerVariables("QUERY_STRING")
End Function

Function IsValidMemName(memname)
    Dim i, c
    IsValidMemName = True
    If Not (3<= Len(memname) And Len(memname)<= 16) Then
        IsValidMemName = False
        Exit Function
    End If
    For i = 1 To Len(memname)
        c = Mid(memname, i, 1)
        If InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-", c) <= 0 And Not IsNumeric(c) Then
            IsValidMemName = False
            Exit Function
        End If
    Next
End Function

Function CheckStr(Strer, Num)
    Dim Shield, w
    If Strer = "" Or IsNull(Strer) Then Exit Function
    Select Case Num
        Case 1
            If IsNumeric(Strer) = 0 Then
                Response.Write "操作错误"
                Response.End
            End If
            Strer = Int(Strer)
    End Select
    CheckStr = Strer
End Function

sub sunlynetpy()
	sunlynet=replace(sunlynet,":","-")
	sunlynet=replace(sunlynet,"：","-")
	sunlynet=replace(sunlynet,"“","")
	sunlynet=replace(sunlynet,"”","")
	sunlynet=replace(sunlynet,"？","-")
	sunlynet=replace(sunlynet," ","-")
end sub

' ====== 用到
Function StrReplace(Str)
    If Str = "" Or IsNull(Str) Then
        StrReplace = ""
        Exit Function
    Else
        StrReplace = Replace(Str, " ", "&nbsp;")
        StrReplace = Replace(StrReplace, Chr(13), "&lt;br&gt;")
        StrReplace = Replace(StrReplace, "<", "&lt;")
        StrReplace = Replace(StrReplace, ">", "&gt;")
    End If
End Function

Function IsValidEmail(email)
    Dim names, Name, i, c
    IsValidEmail = True
    names = Split(email, "@")
    If UBound(names) <> 1 Then
        IsValidEmail = False
        Exit Function
    End If
    For Each Name in names
        If Len(Name) <= 0 Then
            IsValidEmail = False
            Exit Function
        End If
        For i = 1 To Len(Name)
            c = Mid(Name, i, 1)
            If InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-.", c) <= 0 And Not IsNumeric(c) Then
                IsValidEmail = False
                Exit Function
            End If
        Next
        If Left(Name, 1) = "." Or Right(Name, 1) = "." Then
            IsValidEmail = False
            Exit Function
        End If
    Next
    If InStr(names(1), ".") <= 0 Then
        IsValidEmail = False
        Exit Function
    End If
    i = Len(names(1)) - InStrRev(names(1), ".")
    If i <> 2 And i <> 3 Then
        IsValidEmail = False
        Exit Function
    End If
    If InStr(email, "..") > 0 Then
        IsValidEmail = False
    End If
End Function

Function Checkfile(Filepath)
    Dim Fso, Path
    Set Fso = Server.CreateObject("Scripting.Filesystemobject")
    If InStr(Filepath, ":") <> 0 Then Path = Filepath Else Path = Server.Mappath(Filepath)
    If Fso.FileExists(Path) Then Checkfile = True Else Checkfile = False
    Set Fso = Nothing
End Function

Function StrLeft(Str, StrLen)
    Dim L, T, I, C
    If Str = "" Then
        StrLeft = ""
        Exit Function
    End If
    Str = Replace(Replace(Replace(Replace(Str, "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
    L = Len(Str)
    T = 0
    For i = 1 To L
        C = Abs(AscW(Mid(Str, i, 1)))
        If C>255 Then
            T = T + 2
        Else
            T = T + 1
        End If
        If T>= StrLen Then
            StrLeft = Left(Str, i) & "...."
            Exit For
        Else
            StrLeft = Str
        End If
    Next
    StrLeft = Replace(Replace(Replace(Replace(StrLeft, " ", "&nbsp;"), Chr(34), "&quot;"), ">", "&gt;"), "<", "&lt;")
End Function

Function RemoveHTML(strText)
	Dim RegEx
	Set RegEx = New RegExp
	RegEx.Pattern = "<[^>]*>"
	RegEx.Global = True
	RemoveHTML = RegEx.Replace(strText, "")
End Function

Function FormatDate(DateAndTime, Format)
	If Not IsDate(DateAndTime) Then Exit Function
	If IsNul(Format) Then
		FormatDate = DateAndTime
	Else
		If IsNumeric(Format) Then
			FormatDate = FormatDate2(DateAndTime, Format)
		Else
			FormatDate = FormatDate1(DateAndTime, Format)
		End If
	End If
End Function

Function FormatDate1(DateAndTime, Format)
    On Error Resume Next
    Dim yy, y, m, d, h, mi, s, strDateTime
    FormatDate = DateAndTime
    If IsNumeric(Format) Then Exit Function
	Select Case Format
		Case "yy"
			strDateTime = CStr(Year(DateAndTime))
		Case "y"
			strDateTime = Mid(CStr(Year(DateAndTime)), 3)
		Case "mon"
			m = CStr(Month(DateAndTime))
			If Len(m) = 1 Then m = "0" & m
			strDateTime = m
		Case "mo"
			strDateTime = CStr(Month(DateAndTime))
		Case "dd"
			d = CStr(Day(DateAndTime))
			If Len(d) = 1 Then d = "0" & d
			strDateTime = d
		Case "d"
			strDateTime = CStr(Day(DateAndTime))
		Case "hh"
			h = CStr(Hour(DateAndTime))
			If Len(h) = 1 Then h = "0" & h
			strDateTime = h
		Case "h"
			strDateTime = CStr(Hour(DateAndTime))
		Case "min"
			min = CStr(Minute(DateAndTime))
			If Len(min) = 1 Then min = "0" & min
			strDateTime = min
		Case "mi"
			strDateTime = CStr(Minute(DateAndTime))
		Case "ss"
			s = CStr(Second(DateAndTime))
			If Len(s) = 1 Then s = "0" & s
			strDateTime = s
		Case "s"
			strDateTime = CStr(Second(DateAndTime))
		Case Else
			strDateTime = DateAndTime
	End Select
	FormatDate1 = strDateTime
End Function

Function FormatDate2(DateAndTime, Format)
    On Error Resume Next
    Dim yy, y, m, d, h, mi, s, strDateTime
    FormatDate = DateAndTime
    If Not IsNumeric(Format) Then Exit Function
    yy = CStr(Year(DateAndTime))
    y = Mid(CStr(Year(DateAndTime)), 3)
    m = CStr(Month(DateAndTime))
    If Len(m) = 1 Then m = "0" & m
    d = CStr(Day(DateAndTime))
    If Len(d) = 1 Then d = "0" & d
    h = CStr(Hour(DateAndTime))
    If Len(h) = 1 Then h = "0" & h
    mi = CStr(Minute(DateAndTime))
    If Len(mi) = 1 Then mi = "0" & mi
    s = CStr(Second(DateAndTime))
    If Len(s) = 1 Then s = "0" & s
    Select Case Format
        Case "1"
            strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
        Case "2"
            strDateTime = yy & m & d & h & mi & s
        Case "3"
            strDateTime = yy & m & d & h & mi
        Case "4"
            strDateTime = yy & "年" & m & "月" & d & "日"
        Case "5"
            strDateTime = m & "-" & d
        Case "6"
            strDateTime = m & "/" & d
        Case "7"
            strDateTime = m & "月" & d & "日"
        Case "8"
            strDateTime = y & "年" & m & "月"
        Case "9"
            strDateTime = y & "-" & m
        Case "10"
            strDateTime = y & "/" & m
        Case "11"
            strDateTime = y & "-" & m & "-" & d
        Case "12"
            strDateTime = y & "/" & m & "/" & d
        Case "13"
            strDateTime = yy & "." & m & "." & d
        Case "14"
            strDateTime = yy & "-" & m & "-" & d
        Case Else
            strDateTime = DateAndTime
    End Select
    FormatDate2 = strDateTime
End Function

Public SiteTitle, SiteUrl, ComName, Address, Telephone, Mobile, Contacts, Fax, Email, Keywords, Descriptions, IcpNumber, Stats, SiteLogo, SiteDetail, Ico, SiteStatus, Explain
Public MobileTitle, MobileKeywords, MobileDescriptions, MobileLogo, MobileTelephone, MobileJump
Public TemplatePath, mTemplatePath
Public AboutDir, NewsDir, ProDir, CaseDir, DownDir, JobDir, AboutPc, NewsPc, ProPc, CasePc, DownPc, JobPc, AboutMc, NewsMc, ProMc, CaseMc, DownMc, JobMc

Sub SiteInfo()
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select top 1 * from keyicms_Site"
    rs.Open sql, conn, 1, 1
    SiteTitle = rs("SiteTitle")
    Keywords = rs("Keywords")
    Descriptions = rs("Descriptions")
	Stats = rs("Stats")
    SiteUrl = rs("SiteUrl")
    ComName = rs("ComName")
    Address = rs("Address")
    Telephone = rs("Telephone")
	Mobile = rs("Mobile")
    Contacts = rs("Contacts")
    Fax = rs("Fax")
    Email = rs("Email")
    IcpNumber = rs("IcpNumber")
    SiteLogo = rs("SiteLogo")
    SiteDetail = rs("SiteDetail")
    SiteStatus = rs("SiteStatus")
    Explain = rs("Explain")
    Ico = rs("Ico")
    rs.Close : Set rs = Nothing

    Set rs = server.CreateObject("adodb.recordset")
    sql = "select top 1 * from keyicms_MobileSite"
    rs.Open sql, conn, 1, 1
	If mStatus Then
		MobileTitle = rs("MobileTitle")
		MobileKeywords = rs("MobileKeywords")
		MobileDescriptions = rs("MobileDescriptions")
		MobileLogo = rs("MobileLogo")
		MobileTelephone = rs("MobileTelephone")
		MobileJump = rs("MobileJump")
	End If
    rs.Close : Set rs = Nothing

	If not Conn.execute("select * from keyicms_Template where UseFlag and SortName='PC'").eof Then
		TemplatePath  = Conn.execute("select top 1 * from keyicms_Template where UseFlag and SortName='PC'")("TemplatePath")
	End If
	
	If not Conn.execute("select * from keyicms_Template where UseFlag and SortName='Mobile'").eof Then
		mTemplatePath = Conn.execute("select top 1 * from keyicms_Template where UseFlag and SortName='Mobile'")("TemplatePath")
	End If

	AboutDir = Conn.execute("select * from keyicms_System where DataForm=1")("SystemDir")
	NewsDir  = Conn.execute("select * from keyicms_System where DataForm=2")("SystemDir")
	ProDir   = Conn.execute("select * from keyicms_System where DataForm=3")("SystemDir")
	CaseDir  = Conn.execute("select * from keyicms_System where DataForm=4")("SystemDir")
	DownDir  = Conn.execute("select * from keyicms_System where DataForm=5")("SystemDir")
	JobDir   = Conn.execute("select * from keyicms_System where DataForm=6")("SystemDir")
End Sub

'文字转换成拼音
'========================================
Function getEnglish(Str)
    getEnglish = Str
	getEnglish = Replace(getEnglish, "狗", " dog ")
	getEnglish = Replace(getEnglish, "科蚁", " ke01 ")
	getEnglish = Replace(getEnglish, "中国", " china ")
	getEnglish = Replace(getEnglish, "策划", " plan ")
	getEnglish = Replace(getEnglish, "免费", " free ")
	getEnglish = Replace(getEnglish, "介绍", " intro ")
	getEnglish = Replace(getEnglish, "技巧", " skill ")
	getEnglish = Replace(getEnglish, "生活", " life ")
	getEnglish = Replace(getEnglish, "活动", " activity ")
	getEnglish = Replace(getEnglish, "工具", " tool ")
	getEnglish = Replace(getEnglish, "联盟", " union ")
	getEnglish = Replace(getEnglish, "注册", " register ")
	getEnglish = Replace(getEnglish, "经验", " experience ")
	getEnglish = Replace(getEnglish, "翻译", " translate ")
	getEnglish = Replace(getEnglish, "项目", " item ")
	getEnglish = Replace(getEnglish, "网站", " web-site ")
	getEnglish = Replace(getEnglish, "英语", " english ")
	getEnglish = Replace(getEnglish, "英文", " english ")
	getEnglish = Replace(getEnglish, "雅虎", " yahoo ")
	getEnglish = Replace(getEnglish, "新浪", " sina ")
	getEnglish = Replace(getEnglish, "交易", " trade ")
	getEnglish = Replace(getEnglish, "网店", " b2c ")
	getEnglish = Replace(getEnglish, "升级", " update ")
	getEnglish = Replace(getEnglish, "杂志", " magazine ")
	getEnglish = Replace(getEnglish, "空间", " space ")
	getEnglish = Replace(getEnglish, "爱情", " love ")
	getEnglish = Replace(getEnglish, "朋友", " friend ")
	getEnglish = Replace(getEnglish, "友情", " friend ")
	getEnglish = Replace(getEnglish, "链接", " like ")
	getEnglish = Replace(getEnglish, "标签", " label ")
	getEnglish = Replace(getEnglish, "运行", " running ")
	getEnglish = Replace(getEnglish, "管理", " manager ")
	getEnglish = Replace(getEnglish, "管理", " manage ")
	getEnglish = Replace(getEnglish, "页面", " page ")
	getEnglish = Replace(getEnglish, "模板", " template ")
	getEnglish = Replace(getEnglish, "游戏", " game ")
	getEnglish = Replace(getEnglish, "论坛", " forum ")
	getEnglish = Replace(getEnglish, "新闻", " news ")
	getEnglish = Replace(getEnglish, "音乐", " music ")
	getEnglish = Replace(getEnglish, "帮助", " help ")
	getEnglish = Replace(getEnglish, "优化", " optimize ")
	getEnglish = Replace(getEnglish, "软件", " soft ")
	getEnglish = Replace(getEnglish, "教程", " tech ")
	getEnglish = Replace(getEnglish, "下载", " download ")
	getEnglish = Replace(getEnglish, "搜索", " search ")
	getEnglish = Replace(getEnglish, "引擎", " engine ")
	getEnglish = Replace(getEnglish, "蜘蛛", " spider ")
	getEnglish = Replace(getEnglish, "日志", " log ")
	getEnglish = Replace(getEnglish, "博客", " blog ")
	getEnglish = Replace(getEnglish, "百度", " baidu ")
	getEnglish = Replace(getEnglish, "谷歌", " google ")
	getEnglish = Replace(getEnglish, "邮箱", " mailbox ")
	getEnglish = Replace(getEnglish, "邮件", " mail ")
	getEnglish = Replace(getEnglish, "域名", " domain ")
	getEnglish = Replace(getEnglish, "测试", " test")
	getEnglish = Replace(getEnglish, "演示", " demo ")
	getEnglish = Replace(getEnglish, "音乐", " music ")
	getEnglish = Replace(getEnglish, "笑话", " joke ")
	getEnglish = Replace(getEnglish, "产品", " product ")
	getEnglish = Replace(getEnglish, "留言", " message ")
	getEnglish = Replace(getEnglish, "反馈", " freedback ")
	getEnglish = Replace(getEnglish, "评论", " comment ")
	getEnglish = Replace(getEnglish, "推荐", " commend ")
	getEnglish = Replace(getEnglish, "共享", " share ")
	getEnglish = Replace(getEnglish, "资源", " resource ")
	getEnglish = Replace(getEnglish, "插件", " plugins ")
	getEnglish = Replace(getEnglish, "本本", " notebook ")
	getEnglish = Replace(getEnglish, "电脑", " computer ")
	getEnglish = Replace(getEnglish, "系统", " system ")
	getEnglish = Replace(getEnglish, "学校", " school ")
	getEnglish = Replace(getEnglish, "工作", " job ")
	getEnglish = Replace(getEnglish, "信息", " info ")
	getEnglish = Replace(getEnglish, "娱乐", " ent ")
	getEnglish = Replace(getEnglish, "汽车", " car ")
	getEnglish = Replace(getEnglish, "手机", " mobile ")
	getEnglish = Replace(getEnglish, "网络", " network ")
	getEnglish = Replace(getEnglish, "老板", " boss ")
	getEnglish = Replace(getEnglish, "电视", " tv ")
	getEnglish = Replace(getEnglish, "电影", " movie ")
	getEnglish = Replace(getEnglish, "支付宝", " alipay ")
End Function

Function Pinyin(chinese)
	chinese = replace(chinese, "/", ""): chinese = replace(chinese, "\", "")
	chinese = replace(chinese, "*", ""): chinese = replace(chinese, "]", "")
	chinese = replace(chinese, "[", ""): chinese = replace(chinese, "}", "")
	chinese = replace(chinese, "{", ""): chinese = replace(chinese, "'", "")
	chinese = getEnglish(chinese)
	dim pinyinstr, istr, iIsCn, IsCn
	dim pinyinconn, rs, i, x
	on error resume next
	set pinyinconn = server.createobject("Adodb.connection")
	pinyinconn.open "Provider=Microsoft.Jet.OLEdb.4.0;Data Source=" & server.mappath("pinyin.asp")
	if err then pinyin = "" : set pinyinconn = nothing : exit function
	IsCn = true
	for i = 1 to len(chinese)
		iIsCn = IsCn '获取上次是不是中文的值
		istr = Mid(chinese, i, 1)
		Session.Codepage = 936 ' 修改成 936
'		x = Asc(istr) + 65536
		x = Asc(istr)
		Session.Codepage = 65001 ' 用完改回来	
		
		if (x >= 65 and x <= 90) or (x >= 97 and x <= 122) or (x >= 48 and x <= 57) or istr = " " then
			IsCn = False '这些是英文、数字(保留字符)
			if istr = " " then istr = "-"
		else
			set rs = pinyinconn.execute("select top 1 pinyin From pinyin where content like '%" & istr & "%';")
			if not rs.eof then
				'istr = lcase(rs(0)): IsCn = True '小写
				istr = rs(0): IsCn = True '首字母大写
			else
				IsCn = false
				if istr = " " then istr = "-" else istr = "" '将空格转换成"－"
			end if
			rs.close : set rs = nothing
		end if
		if iIsCn = IsCn then pinyinstr = pinyinstr & istr else pinyinstr = pinyinstr & "-" & istr
		pinyinstr = replace(pinyinstr, "--", "-")
		pinyinstr = replace(pinyinstr, "__", "_")
	next

	if right(pinyinstr, 1) = "-" then pinyinstr = left(pinyinstr, len(pinyinstr) - 1)
	if right(pinyinstr, 1) = "_" then pinyinstr = left(pinyinstr, len(pinyinstr) - 1)
	if left(pinyinstr, 1) = "-" then pinyinstr = right(pinyinstr, len(pinyinstr) - 1)
	if left(pinyinstr, 1) = "_" then pinyinstr = right(pinyinstr, len(pinyinstr) - 1)
	pinyinconn.close : set pinyinconn = nothing
	Pinyin = trim(pinyinstr)
End Function

'删除静态页面
Sub DoDelHtml(htmlname)
    On Error Resume Next
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
    ServerMap = server.MapPath(htmlname)
    FSO.DeleteFile(ServerMap)
    Set FSO = Nothing
End Sub

'新建文件夹
Sub CreateFolder(strFolder)
	'首选判断要建立的文件夹是否已经存在
	Dim strTestFolder, objFSO
	strTestFolder = Server.Mappath(strFolder)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	' 检查文件夹是否存在
	If not objFSO.FolderExists(strTestFolder) Then
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		' 建立文件夹
		objFSO.CreateFolder(strTestFolder)
		Set objFSO = Nothing
     End If
End Sub

'修改文件夹名称
Sub fldrename(nowfld, newfld)
	nowfld2 = server.mappath(nowfld) 
	newfld2 = server.mappath(newfld) 
	Set fso = CreateObject("Scripting.FileSystemObject")

	if fso.FolderExists(nowfld2) then
		if not fso.FolderExists(newfld2) then 
			fso.CopyFolder nowfld2, newfld2
			fso.DeleteFolder(nowfld2) 
		end if 
	else
		fso.CreateFolder(newfld2)
	end if 
	set fso = nothing 
End Sub

'删除文件夹
Sub DelFileFolder(strFolder)
	Dbpath=server.mappath(strFolder)
	dim fso,strRubbishFile
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	if fso.FolderExists(Dbpath) then    '判断是否有这文件夹
		fso.DeleteFolder(Dbpath)        '删除文件夹
	end if
End Sub

Function TimeToName()
	yy = CStr(Year(now()))
	mm = CStr(Month(now()))
	dd = CStr(Day(now()))
	hh = CStr(Hour(now()))
	mi = CStr(Minute(now()))
	ss = CStr(Second(now()))

    If Len(mm) = 1 Then mm = "0" & mm
    If Len(dd) = 1 Then dd = "0" & dd
    If Len(hh) = 1 Then hh = "0" & hh
    If Len(mi) = 1 Then mi = "0" & mi
    If Len(ss) = 1 Then ss = "0" & ss

	TimeToName = yy & mm & dd & hh & mi & ss
End Function

Function SweetAlert(SweetType, title, text, showButton, action)
	response.Write "<!doctype html>" & vbCrLf
	response.Write "<link rel='stylesheet' type='text/css' href='SweetAlert/sweet-alert.css'>" & vbCrLf
	response.Write "<script src='SweetAlert/sweet-alert.min.js'></script>" & vbCrLf
	response.Write "<script>window.onload = function(){swal({title: '"&title&"',text: '"&text&"',type: '"&SweetType&"',showCancelButton: "&showButton&",confirmButtonColor: '#DD6B55',confirmButtonText: '确定'},function(){"&action&"});}</script>"
End Function
%>