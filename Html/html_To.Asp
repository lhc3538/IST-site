<%
'检查普通列表或带翻页列表
Function HtmlTo(replace_code, Htmla, Page, Ke01, SortPath)
	If Htmla = "News" Then
		replace_code = NewsList(replace_code, Page, Ke01, SortPath)
	Else
		replace_code = NewsList(replace_code, "", Ke01, "")
	End If
	If Htmla = "Case" Then
		replace_code = CaseList(replace_code, Page, Ke01, SortPath)
	Else
		replace_code = CaseList(replace_code, "", Ke01, "")
	End If
	If Htmla = "Product" Then
		replace_code = ProductList(replace_code, Page, Ke01, SortPath)
	Else
		replace_code = ProductList(replace_code, "", Ke01, "")
	End If
	If Htmla = "Job" Then
		replace_code = JobList(replace_code, Page, Ke01)
	Else
		replace_code = JobList(replace_code, "", Ke01)
	End If
	If Htmla = "Download" Then
		replace_code = DownList(replace_code, Page, Ke01, SortPath)
	Else
		replace_code = DownList(replace_code, "", Ke01, "")
	End If
	replace_code = About(replace_code, "", Ke01, "")
	
	replace_code = KE01_If(replace_code)
	HtmlTo = replace_code
End Function
%>