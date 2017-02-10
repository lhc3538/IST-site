<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/System_Language.asp" -->
<!--#include file="../Include/Version.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
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
	If InStr(CompanyAdminPurview, "|13,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> 当前位置</li>
                <li>企业应用</li>
                <%
				If SortPath<>"" Then
					Response.Write "<li>留言信息列表 -> <a href='Ky_MessageList.Asp'>全部</a></li>"
				Else
					Response.Write "<li class='active'>留言信息列表 -> 全部</li>"
				End If
				%>
            </ol>	
        </div>	
        <div class="row">
            <div class="col-lg-12">
                <div class="panel panel-default">
                    <div data-original-title="" class="panel-heading">
                        <h2><i class="fa fa-user"></i><span class="break"></span>留言信息列表</h2>
                    </div>
                    
                    <div class="panel-body">
                        <div id="DataTables_Table_0_wrapper" class="dataTables_wrapper form-inline no-footer">
                        	<input type="hidden" name="Result" value="Message">
                            <table id="datatable" class="table table-striped table-bordered bootstrap-datatable" cellspacing="0" width="100%">
                            <thead>
                            <tr role="row">
                                <th>ID</th>
                                <th>留言主题</th>
                                <th>留言人</th>
                                <th>留言时间</th>
                                <th>操作</th>
                                <th>选择</th>
                            </tr>
                            </thead>
                            <tbody>
								<%Call MessageList()%>
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
</html>
<%
Function MessageList()
	datafrom = "Keyicms_Message"
	taxis = "Order by ID desc"
	sql = "select * from ["& datafrom &"] "&taxis
	
	Set rs = server.CreateObject("adodb.recordset")
	rs.Open sql, conn, 1, 2
	do until rs.EOF
		Response.Write "<tr role='row'>" & vbCrLf
		Response.Write "<td>"&rs("ID")&"</td>" & vbCrLf
		If StrLen((rs("MesName")))>40 Then
			MesName = StrLeft(rs("MesName"), 37)
		Else
			MesName = rs("MesName")
		End If
		Response.Write "<td title='联系方式："&rs("Telephone")&"'>"&MesName&SeeFlag&"</td>" & vbCrLf
		Response.Write "<td>"&rs("Linkman")&"</td>" & vbCrLf
		Response.Write "<td>"&rs("AddTime")&"</td>" & vbCrLf
		Response.Write "<td>" & vbCrLf
		Response.Write "<a title='查看' href='Ky_MessageAdd.Asp?Result=Modify&ID="&rs("ID")&"' class='btn btn-success'><i class='fa fa-search-plus'></i></a>"
		Response.Write "<a title='删除' href='JavaScript:void(0)' class='btn btn-danger' onclick='test("&rs("ID")&", ""Message"")'><i class='fa fa-trash-o'></i></a>"
		Response.Write "</td>" & vbCrLf
		Response.Write "<td><input name='SelectID' type='checkbox' value='"&rs("ID")&"'></td>"
		Response.Write "</tr>" & vbCrLf
		rs.movenext
	loop
End Function
%>