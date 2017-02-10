<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/System_Language.asp" -->
<!--#include file="../Include/Version.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
'*******************************************************
'软件名称：科蚁企业网站内容管理系统（KEYICMS）
'软件开发：成都神笔天成网络科技有限公司
'网   址：http://www.keyicms.com
'本信息不会影响您网站的正常使用，无论免费用户或是收费用户请保留这里的信息.
'神笔(16)天成网络享有软件著作权，如果您擅自删除这里的版权信息,我们将追究法律责任.
'未经公司同意禁止更改以下代码.
'*******************************************************
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
	If InStr(CompanyAdminPurview, "|11,") = 0 Then
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
					Response.Write "<li>广告模块</li>"
					TextPath(SortID)
				Else
					Response.Write "<li class='active'>广告模块</li>"
				End If
				%>
            </ol>	
        </div>
        <div class="row">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>图文管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_AdsEdit.Asp?Result=Add">  
                            <i class="fa fa-inbox"></i>
                            <div class="title">添加图文信息</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_Ads.Asp">  
                            <i class="fa fa-book"></i>
                            <div class="title">图文管理</div>
                        </a>														
                        <a class="ShortCut3" href="Ky_AdsSort.Asp">  
                            <i class="fa fa-sitemap"></i>
                            <div class="title">图文类别</div>
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
				Set rs=server.CreateObject("adodb.recordset")
				sql="Select * From keyicms_AdsSort where ID="&request.QueryString("id")
				rs.open sql,conn,1,1
				SortPath = rs("SortPath")
				conn.execute("delete from keyicms_AdsSort where Instr(SortPath,'"&SortPath&"')>0")
				conn.execute("delete from keyicms_Ads where Instr(SortPath,'"&SortPath&"')>0")
				Call SweetAlert("success", "操作成功", "成功删除本类、子类及所有下属信息条目！", "false", "location.replace('Ky_AdsSort.Asp')")
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
				<h2><i class="fa fa-user"></i><span class="break"></span>管理图文类别 &nbsp;&nbsp;&nbsp;&nbsp; <a href="Ky_AdsSort.Asp?Action=Add&ParentID=0" class="btn btn-danger btn-sm">添加一级分类</a></h2>
			</div>
			
			<div class="panel-body">
                <table class="table table-striped table-bordered bootstrap-datatable" cellspacing="0" width="100%">
                <thead>
                <tr role="row">
                    <th>ID</th>
                    <th>类别名称</th>
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
	ParentID = request.QueryString("ParentID")
	addFolderForm ParentID
end Function

Function addFolderForm(ParentID)
	Dim ParentPath, SortTextPath, rs, sql
	If ParentID = 0 Then
		ParentPath = "0,"
		SortTextPath = ""
	Else
		Set rs = server.CreateObject("adodb.recordset")
		sql = "Select * From keyicms_AdsSort where ID="&ParentID
		rs.Open sql, conn, 1, 1
		ParentPath = rs("SortPath")
	End If
%>
<div class="row">
    <div class="col-md-12">
        <div class="panel panel-default">
            <div class="panel-heading">
                <h2><i class="fa fa-gears"></i><span class="break"></span><strong>【添加图文类别】| 根类 → <% if ParentID<>0 then TextPath(ParentID)%></strong></h2>
            </div>
            <div class="panel-body">
            <form name="FolderForm" method="post" action="Ky_AdsSort.Asp?Action=Save&From=Add" class="form-horizontal">
                <div class="form-group">
                    <label class="col-md-2 control-label">类别名称：</label>
                    <div class="col-md-4">
                        <input name="SortName" type="text" id="SortName" class="form-control" maxlength="100">
                        <label class="checkbox-inline" for="ViewFlag">
                            <input type="checkbox" id="ViewFlag" name="ViewFlag" value="1"> 是否生效
                        </label>                           
                    </div>
                </div> 
                <input name="ParentID" type="hidden" id="ParentID" value="<%=ParentID%>">
                <input name="ParentPath" type="hidden" id="ParentPath" value="<%=ParentPath%>" >
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

  ID=request.QueryString("ID")
  editFolderForm (ID)
end function

Function editFolderForm(ID)
	Dim SortName, ViewFlag, ParentID, SortPath, rs, sql
	Set rs = server.CreateObject("adodb.recordset")
	sql = "Select * From keyicms_AdsSort where ID="&ID
	rs.Open sql, conn, 1, 1
	SortName = rs("SortName")
	ViewFlag = rs("ViewFlag")
	ParentID = rs("ParentID")
	SortPath = rs("SortPath")
%>
<div class="row">
    <div class="col-md-12">
        <div class="panel panel-default">
            <div class="panel-heading">
                <h2><i class="fa fa-gears"></i><span class="break"></span><strong>【修改图文类别】| 根类 → <% if ParentID<>0 then TextPath(ParentID)%></strong></h2>
            </div>
            <div class="panel-body">
            <form name="FolderForm" method="post" action="Ky_AdsSort.Asp?Action=Save&From=Edit&ID=<%=ID%>" class="form-horizontal">
                <div class="form-group">
                    <label class="col-md-2 control-label" for="text-input">类别名称：</label>
                    <div class="col-md-4">
                        <input name="SortName" type="text" id="SortName" class="form-control" value="<%=SortName%>" maxlength="100">
                        <label class="checkbox-inline" for="ViewFlag">
                            <input type="checkbox" id="ViewFlag" name="ViewFlag" value="1" <%if ViewFlag then response.write ("checked")%>> 是否生效
                        </label>                           
                    </div>
                </div> 
                <input name="ParentID" type="hidden" id="ParentID" value="<%=ParentID%>">
                <input name="SortPath" type="hidden" id="SortPath" value="<%=SortPath%>" >
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

Function TextPath(ID)
	Dim rs,sql,SortTextPath
	Set rs = server.CreateObject("adodb.recordset")
	sql = "Select * From keyicms_AdsSort where ID="&ID
	rs.open sql,conn,1,1
	SortTextPath = rs("SortName")&"&nbsp;→&nbsp;"
	if rs("ParentID") <> 0 then TextPath rs("ParentID")
	response.write(SortTextPath)
End Function

Function saveFolder
    If Len(Trim(request.Form("SortName"))) = 0 Then
		Call SweetAlert("warning", "友情提示", "请填写类别名称！", "false", "history.back(-1)")
        response.End
    End If

	Dim From, Action, rs, sql, SortTextPath
    From = request.QueryString("From")
    Set rs = server.CreateObject("adodb.recordset")
    If From = "Add" Then
        sql = "Select * From keyicms_AdsSort"
        rs.Open sql, conn, 1, 3
        rs.addnew
        Action = "添加广告类别"
        rs("SortPath") = request.Form("ParentPath") & rs("ID") &","
    Else
        sql = "Select * From keyicms_AdsSort where ID="&request.QueryString("ID")
        rs.Open sql, conn, 1, 3
        Action = "修改广告类别"
        rs("SortPath") = request.Form("SortPath")
    End If

    rs("SortName") = request.Form("SortName")
	If request.Form("ViewFlag") = "" Then
		rs("ViewFlag") = 0
	Else
		rs("ViewFlag") = request.Form("ViewFlag")
	End If
    rs("ParentID") = request.Form("ParentID")

    rs.update
	rs.close : set rs=nothing
	Call SweetAlert("success", "操作成功", Action&"成功！", "false", "location.replace('Ky_AdsSort.Asp')")
End Function
%>
<!--#Include file="Keyicms_js.Asp"-->
</body>
</html>
<%
Function Folder(id)
Dim rs,sql,i,ChildCount,FolderType,FolderName,onMouseUp,ListType
Set rs = server.CreateObject("adodb.recordset")
sql = "Select * From keyicms_AdsSort where ParentID="&id&" order by id"
rs.open sql,conn,1,1
if id = 0 and rs.recordcount = 0 then
	response.write "<center>暂无广告分类</center>"
end if
i = 1
while not rs.eof
	ChildCount = conn.execute("select count(*) from keyicms_AdsSort where ParentID="&rs("id"))(0)
	if ChildCount = 0 then
		FolderType = "SortFile"
		FolderName = rs("SortName")
		onMouseUp = ""
	else
		FolderType = "SortFolderOpen"
		onMouseUp = "EndSortChange('a"&rs("id")&"','b"&rs("id")&"');"
		FolderName = rs("SortName")
	end If
	datafrom = "keyicms_AdsSort"
	Response.Write "<tr role='row'>" & vbCrLf
	response.write "<td>"&rs("ID")&"</td>"
	response.write "<td>"&FolderName&"</td>"
	If rs("ViewFlag") Then
		Response.Write "<td><a href=""Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=ViewFlag&Operation=down""><span class='label label-success'>生效</span></a></td>" & vbCrLf
	Else
		Response.Write "<td><a href=""Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=ViewFlag&Operation=up""><span class='label label-default'>未生效</span</a></td>" & vbCrLf
	End If
	Response.Write "<td>"
	Response.Write "<a href='Ky_AdsSort.Asp?Action=Add&ParentID="&rs("id")&"' class='btn btn-primary' title='添加子级'><i class='fa fa-plus'></i> 添加子级</a>"
	Response.Write "<a href='Ky_AdsSort.Asp?Action=Edit&ID="&rs("id")&"' class='btn btn-info' title='修改'><i class='fa fa-edit'></i></a>"
	Response.Write "<a href='Javascript:ConfirmDelSort(""Ky_AdsSort"","&rs("id")&")' class='btn btn-danger' title='删除'><i class='fa fa-trash-o '></i></a>"
	Response.Write "</td>" & vbCrLf
	Response.Write "</tr>" & vbCrLf
	If ChildCount>0 Then
		Call Folderr(rs("id"))
	End If
	rs.movenext
	i = i + 1
wend
rs.close : set rs = nothing
End Function

Function Folderr(id)
Dim rs,sql,i,ChildCount,FolderType,FolderName,onMouseUp,ListType
Set rs = server.CreateObject("adodb.recordset")
sql = "Select * From keyicms_AdsSort where ParentID="&id&" order by id"
rs.open sql,conn,1,1
if id = 0 and rs.recordcount = 0 then
	response.write "<center>暂无广告分类</center>"
end if
i = 1
while not rs.eof
	ChildCount = conn.execute("select count(*) from keyicms_AdsSort where ParentID="&rs("id"))(0)
	if ChildCount = 0 then
		FolderType = "SortFile"
		FolderName = rs("SortName")
		onMouseUp = ""
	else
		FolderType = "SortFolderOpen"
		onMouseUp = "EndSortChange('a"&rs("id")&"','b"&rs("id")&"');"
		FolderName = rs("SortName")
	end If
	datafrom = "keyicms_AdsSort"
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
	Response.Write "<td>"
	Response.Write "<a href='Ky_AdsSort.Asp?Action=Add&ParentID="&rs("id")&"' class='btn btn-primary' title='添加子级'><i class='fa fa-plus'></i> 添加子级</a>"
	Response.Write "<a href='Ky_AdsSort.Asp?Action=Edit&ID="&rs("id")&"' class='btn btn-info' title='修改'><i class='fa fa-edit'></i></a>"
	Response.Write "<a href='Javascript:ConfirmDelSort(""Ky_AdsSort"","&rs("id")&")' class='btn btn-danger' title='删除'><i class='fa fa-trash-o'></i></a>"
	Response.Write "</td>" & vbCrLf
	Response.Write "</tr>" & vbCrLf
	If ChildCount>0 Then
		Call Folderr(rs("id"))
	End If
	rs.movenext
	i = i + 1
wend
rs.close : set rs=nothing
End Function
%>