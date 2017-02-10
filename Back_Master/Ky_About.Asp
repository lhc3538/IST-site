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
</head>

<body>
<!--#Include file="Keyicms_Top.Asp"-->
<div id="wrap">
	<!--#Include file="Keyicms_Meun.Asp"-->
	<%
	If InStr(CompanyAdminPurview, "|1,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row ">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> 当前位置</li>
             	<li>功能模块</li>
                <li class="active">学院信息列表</li>
            </ol>	
        </div>
        <div class="row">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span>学院信息管理</h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_AboutEdit.Asp?Result=Add">  
                            <i class="fa fa-sign-in"></i>
                            <div class="title">添加学院信息</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_About.Asp">  
                            <i class="fa fa-list"></i>
                            <div class="title">学院信息列表</div>
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
                        <h2><i class="fa fa-user"></i><span class="break"></span>学院信息列表</h2>
                    </div>
                    <div class="panel-body">
                        <div id="DataTables_Table_0_wrapper" class="dataTables_wrapper form-inline no-footer">
                        	<input type="hidden" name="Result" value="About">
                            <table id="datatable" class="table table-striped table-bordered bootstrap-datatable" cellspacing="0" width="100%">
                            <thead>
                            <tr role="row">
                                <th>ID</th>
                                <th>状态</th>
                                <%If mStatus Then%><th>手机显示</th><%End If%>
                                <th>标题</th>
                                <th>排序</th>
                                <th>操作</th>
                                <th>选择</th>
                            </tr>
                            </thead>
                            <tbody>
								<%Call AboutList()%>
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
	oTable.fnSort( [ [0,'desc'] ] );
});
</script>
<!--#Include file="Keyicms_js.Asp"-->
</body>

<%
Function AboutList()
	datafrom = "Keyicms_About"
	taxis = "Order by Sequence desc, ID desc"
	sql = "select * from ["& datafrom &"] "&taxis
	
	set rs = server.createobject("adodb.recordset")
	rs.open sql,conn,1,2
	do until rs.eof
		Response.Write "<tr role='row'>" & vbCrLf
		Response.Write "<td>"&rs("ID")&"</td>" & vbCrLf
		If rs("ViewFlag") Then
			Response.Write "<td><a href='Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=ViewFlag&Operation=down'><span class='label label-success'>生效</span></a></td>" & vbCrLf
		Else
			Response.Write "<td><a href='Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=ViewFlag&Operation=up'><span class='label label-default'>未生效</span></a></td>" & vbCrLf
		End If
	If mStatus Then
		If rs("MobileFlag") Then
			Response.Write "<td><a href='Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=MobileFlag&Operation=down'><span class='label label-danger'>显示</span></a></td>" & vbCrLf
		Else
			Response.Write "<td><a href='Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=MobileFlag&Operation=up'><span class='label label-warning'>不显示</span></a></td>" & vbCrLf
		End If
	End If
		Response.Write "<td><a href='Ky_AboutEdit.Asp?Result=Modify&ID="&rs("ID")&"' title="""&rs("AboutName")&""">"&StrLeft(rs("AboutName"),28)&"</a></td>"& vbCrLf
		Response.Write "<td>"&rs("Sequence")&"</td>" & vbCrLf
		Response.Write "<td>"
		Response.Write "<a title='查看' href='../ShowInformation.asp?ID="&rs("ID")&"' class='btn btn-success' target='_blank'><i class='fa fa-search-plus'></i></a>"
		Response.Write "<a title='修改' href='Ky_AboutEdit.Asp?Result=Modify&ID="&rs("ID")&"' class='btn btn-info'><i class='fa fa-edit'></i></a>"
		Response.Write "<a title='删除' href='JavaScript:void(0)' class='btn btn-danger' onclick='test("&rs("ID")&", ""About"")'><i class='fa fa-trash-o'></i></a>"
		Response.Write "<a title='排序' href='Ky_About.Asp?Result=ModifySequence&ID="&rs("ID")&"' class='btn btn-warning'><i class='fa fa-sort'></i></a>"
		Response.Write "</td>" & vbCrLf
		Response.Write "<td><input name='SelectID' type='checkbox' value='"&rs("ID")&"'></td>"
		Response.Write "</tr>" & vbCrLf
		rs.movenext
	loop
End Function

Sub ModifySequence()
	dim rs,sql,ID,AboutName,Sequence
	ID=request.QueryString("ID")
	set rs = server.createobject("adodb.recordset")
	sql="select * from keyicms_About where ID="& ID
	rs.open sql,conn,1,1
	AboutName = rs("AboutName")
	Sequence = rs("Sequence")
	rs.close : set rs = nothing

	response.write "<div class='row'>"
	response.write "    <div class='col-lg-12'>"
	response.write "        <div class='panel panel-default'>"
	response.write "        	<div data-original-title='' class='panel-heading'>"
	response.write "				<h2><i class='fa fa-user'></i><span class='break'></span>学院信息排序管理</h2>"
	response.write "			</div>"
	response.write "			<div class='panel-body'>"
	response.write "				<form action='?Result=SaveSequence' method='post' name='formSequence'>"
	response.write "					<div align='center'>"
	response.write "						ID：<input name='ID' type='text' class='input rounded' value='"&ID&"' readonly />"
	response.write "						学院信息名称：<input name='AboutName' type='text' id='AboutName' class='input rounded' value='"&AboutName&"' readonly />"
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
	dim rs,sql
	set rs = server.createobject("adodb.recordset")
	sql = "select * from keyicms_About where ID="& request.form("ID")
	rs.open sql,conn,1,3
	rs("Sequence") = trim(request.form("Sequence"))
	rs.update
	rs.close : set rs = nothing
	response.redirect "Ky_About.Asp"
End Sub
%>