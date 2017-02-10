<%
'生成首页、网站单页静态方法
Function html_To_Index(FileName, Ke01)
	Call SiteInfo()
	FolderName = FileName
	Template = FileName&".html"
	
	If FileName = "Index" Then
		FilePath = "/Index."&HTMLName 
	Else
		FilePath = "/"&FolderName&"/Index."&HTMLName 
	End If
	If Ke01 = "PC" Then
		replace_code = ReadFromUTF(TemplatePath&"/html/"&Template,"utf-8")
		FilePath = FilePath
		TemplatePath = TemplatePath
	End If
	If Ke01 = "Mobile" Then
		replace_code = ReadFromUTF(mTemplatePath&"/html/"&Template,"utf-8")
		FilePath = "/m"&FilePath
		TemplatePath = mTemplatePath
	End If

	If SiteStatus Then
		replace_code = ReplaceCommon(replace_code, TemplatePath, Ke01)
		replace_code = HtmlTo(replace_code, "", "", Ke01, "")
	Else		
		replace_code = ExplainJS
	End If
	If Ke01 = "PC" Then replace_code = replace_code & ConstJs
	Call WriteToUTF(FilePath, replace_code)
End Function
%>