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
'神笔(1)天成网络享有软件著作权，如果您擅自删除这里的版权信息,我们将追究法律责任.
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
</head>

<body>
<!--#Include file="Keyicms_Top.Asp"-->
<div id="wrap">
	<!--#Include file="Keyicms_Meun.Asp"-->
	<%
	If InStr(CompanyAdminPurview, "|18,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> 当前位置</li>
                <li class="active">网站导航管理</li>
            </ol>	
        </div>	
        <div class="row">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>网站导航管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_NavigationEdit.Asp?Result=Add">  
                            <i class="fa fa-indent"></i>
                            <div class="title">添加导航</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_Navigation.Asp">  
                            <i class="fa fa-th-list"></i>
                            <div class="title">导航管理</div>
                        </a>													
                    </div>
                </div>
            </div>
        </div>
        
		<% if request.QueryString("Result")="ModifySequence" then call ModifySequence() %>
        <% if request.QueryString("Result")="SaveSequence" then call SaveSequence() %>
        
        <div class="row">
            <div class="col-lg-12">
                <div class="panel panel-default">
                    <div data-original-title="" class="panel-heading">
                        <h2><i class="fa fa-user"></i><span class="break"></span>网站导航列表</h2>
                    </div>
                    <div class="panel-body">
                        <div id="DataTables_Table_0_wrapper" class="dataTables_wrapper form-inline no-footer">
                            <table id="datatable" class="table table-striped table-bordered bootstrap-datatable" cellspacing="0" width="100%">
                            <thead>
                            <tr role="row">
                            	<th>位置</th>
                                <th>排序 | 导航名称</th>
                                <th>状态（链接 PC生效<%If mStatus Then%> 手机显示<%End If%>）</th>
                                <th>操作</th>
                            </tr>
                            </thead>
                            <tbody>
								<%Call NavigationList(0, 0) %>
                            </tbody>
                            </table>
                        </div>            
                    </div>
                </div>
            </div>
        </div>
    </div>
	<!--#Include file="Keyicms_End.Asp"-->

</div>
<!--#Include file="Keyicms_js.Asp"-->
</body>
<%
Function NavigationList(ParentID, ClassNum)
Dim rs, sql
datafrom = "Keyicms_Navigation"
Set rs = server.CreateObject("adodb.recordset")
sql = "select * from ["&datafrom&"] where ParentID="&ParentID&" Order by NavType Asc, Sequence Asc, ID Desc"
rs.Open sql, conn, 1, 2
If rs.EOF Then
	Response.Write "<tr><td colspan='4'>暂无相关信息</td></tr>"
Else
	do until rs.EOF
		ChildCount = Conn.execute("select Count(*) from ["&datafrom&"] where ParentID="&rs("ID"))(0)
		Response.Write "<tr>" & vbCrLf
		Response.Write "<td>" & vbCrLf
		If rs("NavType") = 1 then Response.Write "主导航"
		If rs("NavType") = 2 then Response.Write "底部导航"
		Response.Write"</td>" & vbCrLf
		
		Response.Write "<td>"
		Response.Write "<img src='ke01/img/File"&ClassNum&".png' style='padding-right:5px; padding-left:"&ClassNum*40&"px;'/>"&rs("Sequence")&" | "&rs("NavName")&"</td>" & vbCrLf
		Response.Write "<td>" & vbCrLf
		if rs("OutFlag") then
			Response.Write "<span class='label label-danger'>外部</span>" & vbCrLf
		else
			Response.Write "<span class='label label-warning'>内部</span>" & vbCrLf
		end if
		If rs("ViewFlag") Then
			Response.Write "<a href='Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=ViewFlag&Operation=down'><span class='label label-success'>生效</span></a>" & vbCrLf
		Else
			Response.Write "<a href='Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=ViewFlag&Operation=up'><span class='label label-default'>未生效</span></a>" & vbCrLf
		End If
	If mStatus Then
		If rs("MobileFlag") Then
			Response.Write "<a href='Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=MobileFlag&Operation=down'><span class='label label-danger'>显示</span></a>" & vbCrLf
		Else
			Response.Write "<a href='Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=MobileFlag&Operation=up'><span class='label label-warning'>不显示</span></a>" & vbCrLf
		End If
	End If
		response.Write "</td>" & vbCrLf
		Response.Write "<td>" & vbCrLf
		Response.Write "<a title='修改' href='Ky_NavigationEdit.Asp?Result=Modify&ID="&rs("ID")&"' class='btn btn-info'><i class='fa fa-edit '></i></a>"
		Response.Write "<a title='删除' href='JavaScript:void(0)' class='btn btn-danger' onclick='test("&rs("ID")&", ""Navigation"")'><i class='fa fa-trash-o'></i></a>"
		Response.Write "<a title='排序' href='Ky_Navigation.Asp?Result=ModifySequence&ID="&rs("ID")&"' class='btn btn-warning'><i class='fa fa-sort'></i></a>"
		Response.Write "</td>" & vbCrLf
		Response.Write "</tr>" & vbCrLf
		If ChildCount > 0 Then
			Call NavigationList(rs("ID"), ClassNum + 1)
		End If
		rs.movenext
	loop
End If
End Function

Sub ModifySequence()
    Dim rs, sql, ID, NavName, Sequence
    ID = request.QueryString("ID")
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from keyicms_Navigation where ID="& ID
    rs.Open sql, conn, 1, 1
	If rs.eof Then response.Redirect "Ky_Navigation.Asp"
    NavName = rs("NavName")
    Sequence = rs("Sequence")
    rs.Close
    Set rs = Nothing
	
	response.write "<div class='row'>"
	response.write "    <div class='col-lg-12'>"
	response.write "        <div class='panel panel-default'>"
	response.write "        	<div data-original-title='' class='panel-heading'>"
	response.write "				<h2><i class='fa fa-user'></i><span class='break'></span>导航排序管理</h2>"
	response.write "			</div>"
	response.write "			<div class='panel-body'>"
	response.write "				<form action='?Result=SaveSequence' method='post' name='formSequence'>"
	response.write "					<div align='center'>"
	response.write "						ID：<input name='ID' type='text' class='input rounded' value='"&ID&"' readonly />"
	response.write "						导航栏目名称：<input name='NavName' type='text' id='NavName' class='input rounded' value='"&NavName&"' readonly />"
	response.write "						排序号：<input name='Sequence' type='text' class='input rounded' value='"&Sequence&"' maxlength='4' onKeyDown='OnKeyDown()' onchange='OnChange(this, "&Sequence&")' />"
	response.write "						<input name='submitSequence' type='submit' class='button' value='保存'>"
	response.write "					</div>"
	response.write "				</form>"
	response.write "			</div>"
	response.write "		</div>"
	response.write "	</div>"
	response.write "</div>"
End Sub

Sub SaveSequence()
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "select * from keyicms_Navigation where ID="& request.Form("ID")
    rs.Open sql, conn, 1, 3
    rs("Sequence") = request.Form("Sequence")
    rs.update
    rs.Close
    Set rs = Nothing
	Call SweetAlert("success", "友情提示", "修改成功！", "false", "location.replace('Ky_Navigation.Asp');")
End Sub
%>