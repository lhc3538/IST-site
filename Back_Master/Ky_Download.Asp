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
	If InStr(CompanyAdminPurview, "|5,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
	<%
	dim Result,Keyword,SortID,SortPath
	Result = request.QueryString("Result")
	Keyword = request.QueryString("Keyword")
	SortID = request.QueryString("SortID")
	SortPath = request.QueryString("SortPath")
	%>
    <div class="main">
        <div class="row">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> 当前位置</li>
             	<li>功能模块</li>
                <%
				If SortPath<>"" Then
					Response.Write "<li>下载信息列表</li>"
					TextPath(SortID)
				Else
					Response.Write "<li class='active'>下载信息列表</li>"
				End If
				%>
            </ol>	
        </div>	
        <div class="row">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>下载信息管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_DownEdit.Asp?Result=Add">  
                            <i class="fa fa-download"></i>
                            <div class="title">添加下载</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_Download.Asp">  
                            <i class="fa fa-th-list"></i>
                            <div class="title">下载列表</div>
                        </a>	
                        <a class="ShortCut3" href="Ky_DownSort.Asp">  
                            <i class="fa fa-table"></i>
                            <div class="title">下载分类</div>
                        </a>													
                    </div>
                </div>
            </div>
        </div>
        <div class="row">
            <div class="col-lg-12">
                <div class="panel panel-default">
                    <div data-original-title="" class="panel-heading">
                        <h2><i class="fa fa-user"></i><span class="break"></span>下载信息列表</h2>
                        
                    </div>
                    <div class="panel-body">
                        <div id="DataTables_Table_0_wrapper" class="dataTables_wrapper form-inline no-footer">
                        	<input type="hidden" name="Result" value="Download">
                            <table id="datatable" class="table table-striped table-bordered bootstrap-datatable" cellspacing="0" width="100%">
                            <thead>
                            <tr role="row">
                                <th>ID</th>
                                <th>状态</th>
                                <%If mStatus Then%><th>手机显示</th><%End If%>
                                <th>下载标题</th>
                                <th>文件大小</th>
                                <th>操作</th>
                                <th>选择</th>
            				</tr>
                            </thead>
                            <tbody>
								<%Call DownList()%>
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
Function DownList()
	SortID = request.QueryString("SortID")
	SortPath = request.QueryString("SortPath")
	If SortID <> "" Then where = "where instr(SortPath,'"&SortPath&"')>0"
	
	datafrom = "Keyicms_Download"
	taxis = "Order by ID desc"
	sql = "select * from ["& datafrom &"] "&where&" "&taxis
	
	Set rs = server.CreateObject("adodb.recordset")
	rs.Open sql, conn, 1, 2
	do until rs.EOF
		Response.Write "<tr role='row'>" & vbCrLf
		Response.Write "<td>"&rs("ID")&"</td>" & vbCrLf
		If rs("ViewFlag") Then
			Response.Write "<td><a href='Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=ViewFlag&Operation=down'><span class='label label-success'>生效</span></a></td>" & vbCrLf
		Else
			Response.Write "<td><a href='Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=ViewFlag&Operation=up'><span class='label label-default'>未生效</span</a></td>" & vbCrLf
		End If
	If mStatus Then
		If rs("MobileFlag") Then
			Response.Write "<td><a href='Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=MobileFlag&Operation=down'><span class='label label-danger'>显示</span></a></td>" & vbCrLf
		Else
			Response.Write "<td><a href='Conversion.Asp?id="&rs("ID")&"&LX="&datafrom&"&Conver=MobileFlag&Operation=up'><span class='label label-warning'>不显示</span></a></td>" & vbCrLf
		End If
	End If
		if StrLen((rs("DownName")))>40 then
			Response.Write "<td title='"&rs("DownName")&"'><a href='Ky_DownEdit.Asp?Result=Modify&ID="&rs("ID")&"'>"&StrLeft(rs("DownName"),37)&"</a></td>" & vbCrLf
		else
			Response.Write "<td title='"&rs("DownName")&"'><a href='Ky_DownEdit.Asp?Result=Modify&ID="&rs("ID")&"'>"&rs("DownName")&"</a></td>" & vbCrLf
		end if
	
		Response.Write "<td>"&rs("FileSize")&"M</td>" & vbCrLf
		Response.Write "<td>"
		Response.Write "<a title='查看前台页面' href='../ShowDownloadDetail.asp?ID="&rs("ID")&"' class='btn btn-success' target='_blank'><i class='fa fa-search-plus'></i></a>"
		Response.Write "<a title='修改' href='Ky_DownEdit.Asp?Result=Modify&ID="&rs("ID")&"' class='btn btn-info'><i class='fa fa-edit'></i></a>"
		Response.Write "<a title='删除' href='JavaScript:void(0)' class='btn btn-danger' onclick='test("&rs("ID")&", ""Download"")'><i class='fa fa-trash-o'></i></a>"
		Response.Write "</td>" & vbCrLf
		Response.Write "<td><input name='SelectID' type='checkbox' value='"&rs("ID")&"'></td>"
		Response.Write "</tr>" & vbCrLf
		rs.movenext
	loop
End Function

Function TextPath(ID)
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "Select * From keyicms_DownSort where ID="&ID
    rs.Open sql, conn, 1, 1
    TextPath = "<li class='active'>"&rs("SortName")&"</li>"
    If rs("ParentID")<>0 Then TextPath rs("ParentID")
    response.Write(TextPath)
End Function%>