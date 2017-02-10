<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/System_Language.asp" -->
<!--#include file="../Include/Version.Asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
Call SiteInfo()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
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
	If InStr(CompanyAdminPurview, "|6,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row ">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> 当前位置</li>
             	<li>功能模块</li>
                <li class="active">招聘信息列表</li>
            </ol>	
        </div>	
         <div class="row ">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>招聘管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_JobEdit.Asp?Result=Add">  
                            <i class="fa fa-indent"></i>
                            <div class="title">添加招聘信息</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_Job.Asp">  
                            <i class="fa fa-list"></i>
                            <div class="title">招聘信息列表</div>
                        </a>					
                    </div>
                </div>
            </div>
        </div> 

        <div class="row">
            <div class="col-lg-12">
                <div class="panel panel-default">
                    <div data-original-title="" class="panel-heading">
                        <h2><i class="fa fa-user"></i><span class="break"></span>招聘信息列表</h2>
                    </div>
                    <div class="panel-body">
                        <div id="DataTables_Table_0_wrapper" class="dataTables_wrapper form-inline no-footer">
                        	<input type="hidden" name="Result" value="Job">
                            <table id="datatable" class="table table-striped table-bordered bootstrap-datatable" cellspacing="0" width="100%">
                            <thead>
                            <tr role="row">
                                <th>ID</th>
                                <th>状态</th>
                                <%If mStatus Then%><th>手机显示</th><%End If%>
                                <th>招聘标题</th>
                                <th>职位人数</th>
                                <th>发布日期</th>
                                <th>结束日期</th>
                                <th>操作</th>
                                <th>选择</th>
                            </tr>
                            </thead>
                            <tbody>
								<%Call JobList()%>
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
Function JobList()
	datafrom = "Keyicms_Job"
	taxis = "Order by ID desc"
	sql = "select * from ["& datafrom &"] "&taxis
	
	set rs = server.createobject("adodb.recordset")
	rs.open sql,conn,1,2
	do until rs.eof
		Response.Write "<tr role='row'>" & vbCrLf
		Response.Write "<td>"&rs("ID")&"</td>" & vbCrLf
		If rs("ViewFlag") Then
			Response.Write "<td><a href=""Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=ViewFlag&Operation=down""><span class='label label-success'>生效</span></a></td>" & vbCrLf
		Else
			Response.Write "<td><a href=""Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=ViewFlag&Operation=up""><span class='label label-default'>未生效</span</a></td>" & vbCrLf
		End If
	If mStatus Then
		If rs("MobileFlag") Then
			Response.Write "<td><a href='Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=MobileFlag&Operation=down'><span class='label label-danger'>显示</span></a></td>" & vbCrLf
		Else
			Response.Write "<td><a href='Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=MobileFlag&Operation=up'><span class='label label-warning'>不显示</span></a></td>" & vbCrLf
		End If
	End If
		If StrLen((rs("JobName")))>40 Then
			Response.Write "<td title='"&rs("JobName")&"'><a href='Ky_JobEdit.Asp?Result=Modify&ID="&rs("ID")&"''>"&StrLeft(rs("JobName"),37)&"</a></td>" & vbCrLf
		Else
			Response.Write "<td title='"&rs("JobName")&"'><a href='Ky_JobEdit.Asp?Result=Modify&ID="&rs("ID")&"''>"&rs("JobName")&"</a></td>" & vbCrLf
		End If
		Response.Write "<td>"&rs("JobNumber")&"</td>" & vbCrLf
		Response.Write "<td>"&FormatDate2(rs("AddTime"),13)&"</td>" & vbCrLf
		Response.Write "<td>"&FormatDate2(rs("EndDate"),13)&"</td>" & vbCrLf

		Response.Write "<td>"
		Response.Write "<a title='查看前台页面' href='../ShowJobDetail.asp?ID="&rs("ID")&"' class='btn btn-success' target='_blank'><i class='fa fa-search-plus '></i></a>"
		Response.Write "<a title='修改' href='Ky_JobEdit.Asp?Result=Modify&ID="&rs("ID")&"' class='btn btn-info'><i class='fa fa-edit'></i></a>"
		Response.Write "<a title='删除' href='JavaScript:void(0)' class='btn btn-danger' onclick='test("&rs("ID")&", ""Job"")'><i class='fa fa-trash-o'></i></a>"
		Response.Write "</td>" & vbCrLf
		Response.Write "<td><input name='SelectID' type='checkbox' value='"&rs("ID")&"'></td>"
		Response.Write "</tr>" & vbCrLf
		rs.movenext
	loop
End Function
%>