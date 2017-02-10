<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/System_Language.asp" -->
<!--#include file="../Include/Version.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
Call SiteInfo()
%>
<!DOCTYPE html>
<html lang="zh-cn">
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="<%=Keyicms_Lang_Copyright(13)%>">
<meta name="author" content="<%=Keyicms_Lang_Copyright(3)%>">
<title><%=Keyicms_Lang_Copyright(1)%> <%=Str_Soft_Version%></title>
<!--#Include file="Keyicms_css.Asp"-->
<link rel="stylesheet" href="KEditor_keyicms/themes/default/default.css" />
<script charset="utf-8" src="KEditor_keyicms/kindeditor.js"></script>
<script charset="utf-8" src="KEditor_keyicms/lang/zh_CN.js"></script>
<script>
KindEditor.ready(function(K) {
	var editor = K.editor({
		uploadJson : 'KEditor_keyicmsasp/upload_json.asp',
		fileManagerJson : 'KEditor_keyicms/asp/file_manager_json.asp',
		allowFileManager : true
	});
	K('#imagePic').click(function() {
		editor.loadPlugin('image', function() {
			editor.plugin.imageDialog({
				imageUrl : K('#SmallPic').val(),
				clickFn : function(url, title, width, height, border, align) {
					K('#SmallPic').val(url);
					editor.hideDialog();
				}
			});
		});
	});
});
</script>
</head>

<body>
<!--#Include file="Keyicms_Top.Asp"-->
<div id="wrap">
	<!--#Include file="Keyicms_Meun.Asp"-->
	<%
	If InStr(CompanyAdminPurview, "|3,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> 当前位置</li>
                <li class="active">功能模块</li>
                <%
				If SortPath<>"" Then
					Response.Write "<li>图片模块</li>"
					TextPath(SortID)
				Else
					Response.Write "<li class='active'>图片模块</li>"
				End If
				%>
            </ol>	
        </div>
        <div class="row ">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>图片管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_ProductEdit.Asp?Result=Add">  
                            <i class="fa fa-sign-in"></i>
                            <div class="title">添加图片</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_Product.Asp">  
                            <i class="fa fa-th"></i>
                            <div class="title">图片列表</div>
                        </a>	
                        <a class="ShortCut3" href="Ky_ProductSort.Asp">  
                            <i class="fa fa-sitemap"></i>
                            <div class="title">图片分类</div>
                        </a>													
                    </div>
                </div>
            </div>
        </div>
		<%
        Dim Action, FolderName
        Action = request.QueryString("Action")
        Select Case Action
			Case "Add"
				addFolder
				CallFolderView()
			Case "Del"
				Dim rs,sql,SortPath
				SortPath = Conn.execute("Select * From keyicms_ProductSort where ID="&request.QueryString("id"))("SortPath")
				'删除分类对应信息的静态页面
				set rs = server.CreateObject("adodb.recordset")
				rs.open "Select * from keyicms_Product where Instr(SortPath,'"&SortPath&"')>0", conn, 1, 2
				do until rs.eof
					FileName = "/" & ProDir & "/" & rs("ClassSeo") & Separated & rs("ID") & "." & HTMLName
					Call DoDelHtml("/m"&FileName)
					Call DoDelHtml(FileName)
					rs.movenext
				loop
				rs.close : set rs = nothing
				'删除分类文件夹
				Set rs = server.CreateObject("adodb.recordset")
				rs.open "Select * from keyicms_ProductSort where Instr(SortPath,'"&SortPath&"')>0", conn, 1, 2
				do until rs.eof
					FolderName = "/" & ProDir & "/" & rs("FolderName")
					If mStatus Then Call DelFileFolder("/m"&FolderName)
					Call DelFileFolder(FolderName)
					rs.movenext
				loop
				rs.close : set rs = nothing
				'删除分类和信息在数据库里面的数据
				conn.execute("delete from keyicms_ProductSort where Instr(SortPath,'"&SortPath&"')>0")
				conn.execute("delete from keyicms_Product where Instr(SortPath,'"&SortPath&"')>0")
				Call SweetAlert("success", "操作成功", "成功删除本类、子类及所有下属信息条目！", "false", "location.replace('Ky_ProductSort.Asp')")
			Case "Save"
				saveFolder ()
			Case "Edit"
				editFolder ()
				CallFolderView()
			Case "Move"
				moveFolderForm ()
				CallFolderView()
			Case "MoveSave"
				saveMoveFolder ()
			Case Else
				CallFolderView()
        End Select
        %>
    </div>
	<!--#Include file="Keyicms_End.Asp"-->
</div>

<%Function CallFolderView()%>
<div class="row">
	<div class="col-lg-12">
		<div class="panel panel-default">
			<div data-original-title="" class="panel-heading">
				<h2><i class="fa fa-user"></i><span class="break"></span>管理图片类别 &nbsp;&nbsp;&nbsp;&nbsp; <a href="Ky_ProductSort.Asp?Action=Add&ParentID=0" class="btn btn-danger btn-sm">添加一级分类</a></h2>
			</div>
			
			<div class="panel-body">
                <table class="table table-striped table-bordered bootstrap-datatable" cellspacing="0" width="100%">
                <thead>
                <tr role="row">
                    <th>ID</th>
                    <th>分类名称</th>
                    <th>状态</th>
                    <th>操作</th>
                </tr>
                </thead>
                <tbody>
                    <%Folder(0)%>
                </tbody>
                </table>
			</div>
		</div>
	</div>
</div>
<%
End Function
%>       	
<%
Function addFolder()
	Dim ParentID
	ParentID=request.QueryString("ParentID")
	addFolderForm ParentID
end Function

Function addFolderForm(ParentID)
	Dim ParentPath, SortTextPath, rs, sql
	If ParentID = 0 Then
		ParentPath = "0,"
		SortTextPath = ""
	Else
		Set rs = server.CreateObject("adodb.recordset")
		sql = "Select * From keyicms_ProductSort where ID="&ParentID
		rs.Open sql, conn, 1, 1
		ParentPath = rs("SortPath")
	End If
%>
<div class="row">
    <div class="col-md-12">
        <div class="panel panel-default">
            <div class="panel-heading">
                <h2><i class="fa fa-gears"></i><span class="break"></span><strong>【添加图片类别】| 根类 → <% if ParentID<>0 then TextPath(ParentID)%></strong></h2>
            </div>
            <div class="panel-body">
            <form name="FolderForm" method="post" action="Ky_ProductSort.Asp?Action=Save&From=Add" class="form-horizontal">
                <div class="form-group">
                    <label class="col-md-2 control-label">类别名称：</label>
                    <div class="col-md-4">
                        <input name="SortName" type="text" id="SortName" class="form-control" maxlength="100">                         
                    </div>
                </div> 
                <%If EnableSeo = True Then%>
                <div class="form-group">
                    <label class="col-md-2 control-label">类别关键词：</label>
                    <div class="col-md-5">
                    	<input name="SeoKeywords" type="text" id="SeoKeywords" class="form-control" maxlength="250">
                        <span class="help-block">SEO：多个关键词请用“,”半角逗号分开</span>
                    </div>
                </div>
                <div class="form-group">
                    <label class="col-md-2 control-label">类别描述：</label>
                    <div class="col-md-5">
                    	<input name="SeoDescription" type="text" id="SeoDescription" class="form-control" maxlength="250">
                        <span class="help-block">SEO：请填写描述</span>
                    </div>
                </div>
                <%End If%>
                
                <div class="form-group">
                    <div class="col-md-2 control-label"></div>
                    <div class="col-md-10">
                        <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存类别">
                    </div>	
                </div>
            </form>   
            </div>
        </div>
    </div>
</div>
<%
End Function

Function editFolder()
	Dim ID
	ID = request.QueryString("ID")
	editFolderForm (ID)
End Function

Function editFolderForm(ID)
	Dim SortName, ViewFlag, ParentID, SortPath, SeoKeywords, SeoDescription, SmallPic, Content, FolderName, rs, sql
	Set rs = server.CreateObject("adodb.recordset")
	sql = "Select * From keyicms_ProductSort where ID="&ID
	rs.Open sql, conn, 1, 1
	SortName = rs("SortName")
	ViewFlag = rs("ViewFlag")
	ParentID = rs("ParentID")
	SortPath = rs("SortPath")
	SeoKeywords = rs("SeoKeywords")
	SeoDescription = rs("SeoDescription")
	SmallPic = rs("SmallPic")
	Content = rs("Content")
	FolderName = rs("FolderName")
	If SmallPic<>"" or Content<>"" Then
		ShowFlag = True
	Else
		ShowFlag = False
	End If
%>
<div class="row">
    <div class="col-md-12">
        <div class="panel panel-default">
            <div class="panel-heading">
                <h2><i class="fa fa-gears"></i><span class="break"></span><strong>【修改图片类别】| 根类 → <% if ParentID<>0 then TextPath(ParentID)%></strong></h2>
            </div>
            <form name="FolderForm" method="post" action="Ky_ProductSort.Asp?Action=Save&From=Edit&ID=<%=ID%>" class="form-horizontal">
            <div class="panel-body">
                <div class="form-group">
                    <label class="col-md-2 control-label" for="text-input">类别名称：</label>
                    <div class="col-md-4">
                        <input name="SortName" type="text" id="SortName" class="form-control" value="<%=SortName%>" maxlength="100">                         
                    </div>
                </div> 
                <%If EnableSeo = True Then%>
                <div class="form-group">
                    <label class="col-md-2 control-label" for="text-input">类别关键词：</label>
                    <div class="col-md-5">
                    	<input name="SeoKeywords" type="text" id="SeoKeywords" class="form-control" value="<%=SeoKeywords%>" maxlength="250">
                        <span class="help-block">SEO：多个关键词请用“,”半角逗号分开</span>
                    </div>
                </div>
                <div class="form-group">
                    <label class="col-md-2 control-label" for="text-input">类别描述：</label>
                    <div class="col-md-5">
                    	<input name="SeoDescription" type="text" id="SeoDescription" class="form-control" value="<%=SeoDescription%>" maxlength="250">
                        <span class="help-block">SEO：请填写描述</span>
                    </div>
                </div>
                <%End If%>
              
                <div class="form-group">
                    <div class="col-md-2 control-label"></div>
                    <div class="col-md-10">
                        <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存类别">
                    </div>	
                </div>
            </div>
            </form>   
        </div>
    </div>
</div>
<%
End Function

Function TextPath(ID)
  Dim rs,sql,SortTextPath
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From keyicms_ProductSort where ID="&ID
  rs.open sql,conn,1,1
  SortTextPath=rs("SortName")&"&nbsp;→&nbsp;"
  if rs("ParentID")<>0 then TextPath rs("ParentID")
  response.write(SortTextPath)
End Function

Function saveFolder
	Call SiteInfo()
    If Len(Trim(request.Form("SortName"))) = 0 Then
		Call SweetAlert("warning", "友情提示", "请填写类别名称！", "false", "history.back(-1)")
        response.end
    End If

	Dim From, Action, rs, sql, SortTextPath
    From = request.QueryString("From")
	ID = request.QueryString("ID")

	Set rs = server.CreateObject("adodb.recordset")
	rs.Open "Select * From keyicms_ProductSort where FolderName='"&trim(Request.form("FolderName"))&"'", conn, 1, 3
	

	If From = "Add" Then
		FolderName = "/" & ProDir & "/" & trim(Request.form("FolderName"))
		If mStatus Then Call CreateFolder("/m" & FolderName)
		Call CreateFolder(FolderName)
	Else
		NowPath = "/" & ProDir & "/" & Conn.execute("Select * From keyicms_ProductSort where ID="&ID)("FolderName")
		NewPath = "/" & ProDir & "/" & trim(Request.form("FolderName"))
		If mStatus Then Call fldrename("/m" & NowPath, "/m" & NewPath)
		Call fldrename(NowPath, NewPath)
	End If

    Set rs = server.CreateObject("adodb.recordset")
    If From = "Add" Then
        sql = "Select * From keyicms_ProductSort"
        rs.Open sql, conn, 1, 3
        rs.addnew
        Action = "添加图片类别"
        rs("SortPath") = request.Form("ParentPath") & rs("ID") & ","
    Else
        sql = "Select * From keyicms_ProductSort where ID="&request.QueryString("ID")
        rs.Open sql, conn, 1, 3
        Action = "修改图片类别"
        rs("SortPath") = request.Form("SortPath")
    End If

	rs("FolderName") = trim(Request.form("FolderName"))
    rs("SortName") = request.Form("SortName")

	rs("ViewFlag") = 1
	
	if request.Form("ParentID") = "" then
		rs("ParentID") = 0
	else
		rs("ParentID") = request.Form("ParentID")
	end if
	rs("SeoKeywords") = request.Form("SeoKeywords")
	rs("SeoDescription") = request.Form("SeoDescription")
	If request.Form("Show") = "1" Then
		rs("SmallPic") = request.Form("SmallPic")
		rs("Content") = request.Form("Content")
	Else
		rs("SmallPic") = ""
		rs("Content") = ""
	End If
    rs.update
	rs.close : set rs = nothing

	Call SweetAlert("success", "操作成功", Action&"成功！", "false", "location.replace('Ky_ProductSort.Asp')")
End Function
%>
<!--#Include file="Keyicms_js.Asp"-->
</body>
</html>
<%
Function Folder(id)
Dim rs,sql,i,ChildCount,FolderType,FolderName,onMouseUp,ListType
Set rs=server.CreateObject("adodb.recordset")
sql="Select * From keyicms_ProductSort where ParentID="&id&" order by id"
rs.open sql,conn,1,1
if id=0 and rs.recordcount=0 then
	response.write "<center>暂无图片分类</center>"
end if
i = 1
while not rs.eof
	ChildCount=conn.execute("select count(*) from keyicms_ProductSort where ParentID="&rs("id"))(0)
	if ChildCount=0 then
		FolderType="SortFile"
		FolderName=rs("SortName")
		onMouseUp=""
	else
		FolderType="SortFolderOpen"
		onMouseUp="EndSortChange('a"&rs("id")&"','b"&rs("id")&"');"
		FolderName=rs("SortName")
	end If
	datafrom = "keyicms_ProductSort"
	Response.Write "<tr role='row'>" & vbCrLf
	response.write "<td>"&rs("ID")&"&nbsp;"
	response.write "<td>"&FolderName&"&nbsp;"
	If rs("ViewFlag") Then
		Response.Write "<td><a href=""Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=ViewFlag&Operation=down""><span class='label label-success'>生效</span></a></td>" & vbCrLf
	Else
		Response.Write "<td><a href=""Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=ViewFlag&Operation=up""><span class='label label-default'>未生效</span</a></td>" & vbCrLf
	End If
	Response.Write "<td>" & vbCrLf
	Response.Write "<a href='Ky_Product.Asp?SortID="&rs("ID")&"&SortPath="&rs("SortPath")&"' class='btn btn-success'><i class='fa fa-search-plus'></i></a>" & vbCrLf
	Response.Write "<a title='修改' href='Ky_ProductSort.Asp?Action=Edit&ID="&rs("id")&"' class='btn btn-info'><i class='fa fa-edit'></i></a>" & vbCrLf
	Response.Write "</td>" & vbCrLf
	Response.Write "</tr>" & vbCrLf
	If ChildCount>0 Then
		Call Folderr(rs("id"))
	End If
	rs.movenext
	i=i+1
wend
rs.close : set rs = nothing
End Function

Function Folderr(id)
Dim rs,sql,i,ChildCount,FolderType,FolderName,onMouseUp,ListType
Set rs=server.CreateObject("adodb.recordset")
sql="Select * From keyicms_ProductSort where ParentID="&id&" order by id"
rs.open sql,conn,1,1
if id=0 and rs.recordcount=0 then
	response.write "<center>暂无图片分类</center>"
end if
i = 1
while not rs.eof
	ChildCount=conn.execute("select count(*) from keyicms_ProductSort where ParentID="&rs("id"))(0)
	if ChildCount = 0 then
		FolderType = "SortFile"
		FolderName = rs("SortName")
		onMouseUp = ""
	else
		FolderType = "SortFolderOpen"
		onMouseUp = "EndSortChange('a"&rs("id")&"','b"&rs("id")&"');"
		FolderName = rs("SortName")
	end If
	datafrom = "keyicms_ProductSort"
	Response.Write "<tr role='row'>" & vbCrLf
	response.write "<td>"&rs("ID")&"&nbsp;"
	temparr = split(rs("SortPath"),",")
	arrnum = ubound(temparr)
	response.Write "<td><div style=""width:"&(arrnum-1)*25&"px; height:20px; float:left;""></div><div style=""background:url(ke01/img/File"&arrnum-2&".gif) 1px center no-repeat;padding-left:20px; float:left;"">"&FolderName&"&nbsp;"

	If rs("ViewFlag") Then
		Response.Write "<td><a href=""Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=ViewFlag&Operation=down""><span class='label label-success'>生效</span></a></td>" & vbCrLf
	Else
		Response.Write "<td><a href=""Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=ViewFlag&Operation=up""><span class='label label-default'>未生效</span</a></td>" & vbCrLf
	End If

	Response.Write "<td>" & vbCrLf
	Response.Write "<a href='Ky_Product.Asp?SortID="&rs("ID")&"&SortPath="&rs("SortPath")&"' class='btn btn-success'><i class='fa fa-search-plus'></i></a>" & vbCrLf
	Response.Write "<a title='修改' href='Ky_ProductSort.Asp?Action=Edit&ID="&rs("id")&"' class='btn btn-info'><i class='fa fa-edit'></i></a>" & vbCrLf
	Response.Write "</td>" & vbCrLf
	Response.Write "</tr>" & vbCrLf
	If ChildCount>0 Then
		Call Folderr(rs("id"))
	End If
	rs.movenext
	i=i+1
wend
rs.close : set rs = nothing
End Function
%>