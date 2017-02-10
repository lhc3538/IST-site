<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
Dim Result, ID, SelectID
Result = request("Result")
SelectID = request("SelectID")

Call SiteInfo()
Select Case Result
    Case "About"
        If SelectID <> "" Then
			set rs = server.CreateObject("adodb.recordset")
			rs.open "select * from keyicms_About where ID in ("&SelectID&")", conn, 1, 2
			do until rs.eof
				FilePath = "/" & DownDir & "/" & rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName
				Call DoDelHtml("/m"&FilePath)
				Call DoDelHtml(FilePath)
				rs.movenext
			loop
			rs.close : set rs = nothing
			Conn.Execute "delete from keyicms_About where ID in ("&SelectID&")"
		End If
        response.redirect request.servervariables("http_referer")

    Case "Spage"
        If SelectID <> "" Then
			set rs = server.CreateObject("adodb.recordset")
			rs.open "select * from keyicms_Spage where ID in ("&SelectID&")", conn, 1, 2
			do until rs.eof
				FolderName = "/"&rs("FolderName")
				If mStatus Then Call DelFileFolder("/m"&FolderName)
				Call DelFileFolder(FolderName)
				rs.movenext
			loop
			rs.close : set rs = nothing
			Conn.Execute "delete from keyicms_Spage where ID in ("&SelectID&")"
		End If
        response.redirect request.servervariables("http_referer")

	Case "News"
		If SelectID <> "" Then
			set rs = server.CreateObject("adodb.recordset")
			rs.open "select * from keyicms_News where ID in ("&SelectID&")", conn, 1, 2
			do until rs.eof
				FilePath = "/" & NewsDir & "/" & rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName
				Call DoDelHtml("/m"&FilePath)
				Call DoDelHtml(FilePath)
				rs.movenext
			loop
			rs.close : set rs = nothing
			Conn.Execute "delete from keyicms_News where ID in ("&SelectID&")"
		End If
		response.redirect request.servervariables("http_referer")

	Case "Product"
		If SelectID <> "" Then
			set rs = server.CreateObject("adodb.recordset")
			rs.open "select * from keyicms_Product where ID in ("&SelectID&")", conn, 1, 2
			do until rs.eof
				FilePath = "/" & ProDir & "/" & rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName
				Call DoDelHtml("/m"&FilePath)
				Call DoDelHtml(FilePath)
				rs.movenext
			loop
			rs.close : set rs = nothing
			Conn.Execute "delete from keyicms_Product where ID in ("&SelectID&")"
		End If
		response.redirect request.servervariables("http_referer")
		
	Case "Case"
		If SelectID <> "" Then
			set rs = server.CreateObject("adodb.recordset")
			rs.open "select * from keyicms_Case where ID in ("&SelectID&")", conn, 1, 2
			do until rs.eof
				FilePath = "/" & CaseDir & "/" & rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName
				Call DoDelHtml("/m"&FilePath)
				Call DoDelHtml(FilePath)
				rs.movenext
			loop
			rs.close : set rs = nothing
			Conn.Execute "delete from keyicms_Case where ID in ("&SelectID&")"
		End If
		response.redirect request.servervariables("http_referer")
		
	Case "Download"
		If SelectID <> "" Then
			set rs = server.CreateObject("adodb.recordset")
			rs.open "select * from keyicms_Download where ID in ("&SelectID&")", conn, 1, 2
			do until rs.eof
				FilePath = "/" & DownDir & "/" & rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName
				Call DoDelHtml("/m"&FilePath)
				Call DoDelHtml(FilePath)
				rs.movenext
			loop
			rs.close : set rs = nothing
			Conn.Execute "delete from keyicms_Download where ID in ("&SelectID&")"
		End If
		response.redirect request.servervariables("http_referer")
		
	Case "Job"
		If SelectID <> "" Then
			set rs = server.CreateObject("adodb.recordset")
			rs.open "select * from keyicms_Job where ID in ("&SelectID&")", conn, 1, 2
			do until rs.eof
				FilePath = "/" & JobDir & "/" & rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName
				Call DoDelHtml("/m"&FilePath)
				Call DoDelHtml(FilePath)
				rs.movenext
			loop
			rs.close : set rs = nothing
			Conn.Execute "delete from keyicms_Job where ID in ("&SelectID&")"
		End If
		response.redirect request.servervariables("http_referer")

    Case "Navigation"
		If SelectID <> "" Then
			NavPath = Conn.Execute("select * from keyicms_Navigation where ID="&SelectID)("NavPath")
			Conn.Execute "delete from keyicms_Navigation where instr(NavPath, '"&NavPath&"')>0"
		End If
		response.redirect request.servervariables("http_referer")
	
	Case "App"
		If SelectID <> "" Then Conn.Execute "delete from keyicms_App where ID = "&ID
		response.redirect request.servervariables("http_referer")

	Case "Contact"
		if SelectID <> "" Then Conn.execute "delete from keyicms_Contact where ID in ("&SelectID&")"
		response.redirect request.servervariables("http_referer")
	
    Case "FriendLink"
		If SelectID <> "" Then Conn.Execute "delete from Keyicms_FriendLink where ID in ("&SelectID&")"
		response.redirect request.servervariables("http_referer")
		
	Case "Ads"
		If SelectID <> "" Then Conn.Execute "delete from keyicms_Ads where ID in ("&SelectID&")"
		response.redirect request.servervariables("http_referer")
		
	Case "Slide"
		if SelectID <> "" Then Conn.execute "delete from keyicms_Slide where ID in ("&SelectID&")"
		response.redirect request.servervariables("http_referer")

	Case "Message"
		If SelectID <> "" Then Conn.Execute "delete from keyicms_Message where ID in ("&SelectID&")"
		response.redirect request.servervariables("http_referer")
		
	Case "Template"
		If SelectID <> "" Then
			TemplatePath = Conn.execute("Select * from keyicms_Template where ID = "&SelectID)("TemplatePath")
			Call DelFileFolder(TemplatePath)
			Conn.Execute "delete from keyicms_Template where ID = "&SelectID
		End If
		response.redirect request.servervariables("http_referer")

    Case "ManageLog"
		If SelectID <> "" Then Conn.Execute "delete from Keyicms_AdminLog where ID in ("&SelectID&")"
		response.redirect request.servervariables("http_referer")

    Case "Admin"
        If SelectID <> "" Then Conn.Execute "delete from keyicms_Admin where ID in ("&SelectID&")"
        response.redirect request.servervariables("http_referer")
End Select
%>