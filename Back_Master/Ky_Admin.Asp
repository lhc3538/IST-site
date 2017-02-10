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
'神笔(47)天成网络享有软件著作权，如果您擅自删除这里的版权信息,我们将追究法律责任.
'未经公司同意禁止更改以下代码.
'*******************************************************
%>
<!doctype html>
<html>
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
	If InStr(CompanyAdminPurview, "|21,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row ">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i><%=Keyicms_Lang_Copyright(14)%></li>
                <li>管理员管理</li>
                <li class="active">管理员列表</li>
            </ol>	
        </div>
        <div class="row ">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>管理员管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_AdminEdit.Asp?Result=Add">  
                            <i class="fa fa-user"></i>
                            <div class="title">添加管理员</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_Admin.Asp">  
                            <i class="fa fa-users"></i>
                            <div class="title">管理员列表</div>
                        </a>														
                        <a class="ShortCut4" href="Keyicms_ManageLog.Asp">  
                            <i class="fa fa-file-text-o"></i>
                            <div class="title">后台登录日志</div>
                        </a>														
                    </div>
                </div>
            </div>
        </div>
        <div class="row">
            <div class="col-lg-12">
                <div class="panel panel-default">
                    <div data-original-title="" class="panel-heading">
                        <h2><i class="fa fa-user"></i><span class="break"></span>管理员列表</h2>
                    </div>
                    
                    <div class="panel-body">
                        <div id="DataTables_Table_0_wrapper" class="dataTables_wrapper form-inline no-footer">
                        	<input type="hidden" name="Result" value="Admin">
                            <table id="datatable" class="table table-striped table-bordered bootstrap-datatable" cellspacing="0" width="100%">
                            <thead>
                            <tr role="row">
                                <th>ID</th>
                                <th>状态</th>
                                <th>用户名</th>
                                <th>管理权限</th>
                                <th>操作权限</th>
                                <th>操作</th>
                                <th>选择</th>
                            </tr>
                            </thead>
                            <tbody>
								<%Call AdminList()%>
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
<script>
$(document).ready(function(){
	var oTable = $('#datatable').dataTable();
	oTable.fnSort( [ [0,'asc'] ] );
});
</script>
<!--#Include file="Keyicms_js.Asp"-->
</body>
</html>
<%
Function AdminList()
	datafrom = "Keyicms_Admin"
	taxis = "Order by ID desc"
	sql = "select * from ["& datafrom &"] "&taxis
	
	set rs = server.createobject("adodb.recordset")
	rs.open sql,conn,1,2
	do until rs.eof
		Response.Write "<tr role='row'>" & vbCrLf
		Response.Write "<td>"&rs("ID")&"</td>" & vbCrLf
		If rs("working") Then
			Response.Write "<td><span class='label label-success'>生效</span></td>" & vbCrLf
		Else
			Response.Write "<td><span class='label label-default'>未生效</span></td>" & vbCrLf
		End If
	
		Response.Write "<td>"&rs("AdminName")&"</td>"& vbCrLf
		Response.Write "<td>"&rs("UserName")&"</td>" & vbCrLf
		if len(rs("AdminPurview"))>37 then
			Response.Write "<td title='"&rs("Remark")&"'>"&left(rs("AdminPurview"),35)&"…</td>" & vbCrLf
		else
			Response.Write "<td title='"&rs("Remark")&"'>"&rs("AdminPurview")&"</td>" & vbCrLf
		end if 
		Response.Write "<td>" & vbCrLf
		Response.Write "<a title='修改' href='Ky_AdminEdit.Asp?Result=Modify&ID="&rs("ID")&"' class='btn btn-info'><i class='fa fa-edit'></i></a>"
		Response.Write "<a title='删除' href='JavaScript:void(0)' class='btn btn-danger' onclick='test("&rs("ID")&", ""Admin"")'><i class='fa fa-trash-o'></i></a>"
		Response.Write "</td>" & vbCrLf
		Response.Write "<td><input name='SelectID' type='checkbox' value='"&rs("ID")&"'></td>"
		Response.Write "</tr>" & vbCrLf
		rs.movenext
	loop
End Function
%>