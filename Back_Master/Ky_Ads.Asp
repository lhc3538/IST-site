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
'神笔(24)天成网络享有软件著作权，如果您擅自删除这里的版权信息,我们将追究法律责任.
'未经公司同意禁止更改以下代码.
'*******************************************************
dim Result,Keyword,SortID,SortPath,Action
Action = request.QueryString("action")
Result=request.QueryString("Result")
Keyword=request.QueryString("Keyword")
SortID=request.QueryString("SortID")
SortPath=request.QueryString("SortPath")
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
	If InStr(CompanyAdminPurview, "|11,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> 当前位置</li>
                <li>Keyicms</li>
                <%
				If SortPath<>"" Then
					Response.Write "<li>广告信息列表 -> <a href='Ky_Ads.Asp'>全部</a></li>"
					TextPath(SortID)
				Else
					Response.Write "<li class='active'>广告信息列表 -> 全部</li>"
				End If
				%>
            </ol>	
        </div>
        
        <div class="row ">
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
        <div class="row">
            <div class="col-lg-12">
                <div class="panel panel-default">
                    <div data-original-title="" class="panel-heading">
                        <h2><i class="fa fa-user"></i><span class="break"></span>图文信息列表</h2>
                    </div>
                    
                    <div class="panel-body">
                        <div id="DataTables_Table_0_wrapper" class="dataTables_wrapper form-inline no-footer">
                        	<input type="hidden" name="Result" value="Ads">
                            <table id="datatable" class="table table-striped table-bordered bootstrap-datatable" cellspacing="0" width="100%">
                            <thead>
                            <tr role="row">
                                <th>ID</th>
                                <th>图文标题</th>
                                <th>相关图片</th>
                                <th>图文链接</th>
                                <th>操作</th>
                                <th>选择</th>
                            </tr>
                            </thead>
                            <tbody>
								<%Call AdsList()%>
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
Function AdsList()
	SortID = request.QueryString("SortID")
	SortPath = request.QueryString("SortPath")
	If SortID <> "" Then where = "where instr(SortPath,'"&SortPath&"')>0"
	
	datafrom = "Keyicms_Ads"
	taxis = "Order by ID desc"
	sql = "select * from ["& datafrom &"] "&where&" "&taxis
	
	Set rs = server.CreateObject("adodb.recordset")
	rs.Open sql, conn, 1, 2
	do until rs.EOF
		Response.Write "<tr role='row'>" & vbCrLf
		Response.Write "<td>"&rs("ID")&"</td>" & vbCrLf
		Response.Write "<td title='"&rs("AdsName")&"'><a href='Ky_AdsEdit.Asp?Result=Modify&ID="&rs("ID")&"'>"&rs("AdsName")&"</a></td>" & vbCrLf
		If rs("AdsPic") <> "" Then
			Response.Write "<td><a href="&rs("AdsUrl")&" target=_blank><img style=""border:0px"" src="""&rs("AdsPic")&""" width='120' onerror=""this.src='ke01/img/noimg_big.gif'""></a></td>" & vbCrLf
		Else
			Response.Write "<td><a href="&rs("AdsUrl")&" target=_blank>暂无图片</a></td>" & vbCrLf
		End If
		Response.Write "<td>"&rs("AdsUrl")&"</td>" & vbCrLf
		Response.Write "<td>" & vbCrLf
		Response.Write "<a title='修改' href='Ky_AdsEdit.Asp?Result=Modify&ID="&rs("ID")&"' class='btn btn-info'><i class='fa fa-edit'></i></a>" & vbCrLf
		Response.Write "<a title='删除' href='JavaScript:void(0)' class='btn btn-danger' onclick='test("&rs("ID")&", ""Ads"")'><i class='fa fa-trash-o'></i></a>" & vbCrLf
		Response.Write "<td><input name='SelectID' type='checkbox' value='"&rs("ID")&"'></td>"
		Response.Write "</td>" & vbCrLf
		Response.Write "</tr>" & vbCrLf
		rs.movenext
	loop
	rs.close : set rs = nothing
End Function

Function TextPath(ID)
    Dim rs, sql
    Set rs = server.CreateObject("adodb.recordset")
    sql = "Select * From keyicms_AdsSort where ID="&ID
    rs.Open sql, conn, 1, 1
    TextPath = "<li> -> <a href='Ky_Ads.Asp?SortID="&rs("ID")&"&SortPath="&rs("SortPath")&"'>"&rs("SortName")&"</a></li>"
    If rs("ParentID")<>0 Then TextPath rs("ParentID")
    response.Write(TextPath)
End Function
%>