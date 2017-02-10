<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/system_language.asp" -->
<!--#include file="CheckAdmin.Asp"-->
<%
'*******************************************************
'软件名称：科蚁企业网站内容管理系统（KEYICMS）
'软件开发：成都神笔天成网络科技有限公司
'网   址：http://www.keyicms.com
'本信息不会影响您网站的正常使用，无论免费用户或是收费用户请保留这里的信息.
'神笔(28)天成网络享有软件著作权，如果您擅自删除这里的版权信息,我们将追究法律责任.
'未经公司同意禁止更改以下代码.
'*******************************************************
dim Action
dim ID, MapX, MapY, MapLevel, MapLogo, AK
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
<script charset="utf-8" src="KEditor_keyicms/kindeditor-min.js"></script>
<script charset="utf-8" src="KEditor_keyicms/lang/zh_CN.js"></script>
<script charset="utf-8" src="KEditor_keyicms/plugins/code/prettify.js"></script>
<script>
KindEditor.ready(function(K) {
	var editor = K.editor({
		cssPath : 'KEditor_keyicms/plugins/code/prettify.css',
		uploadJson : 'KEditor_keyicms/asp/upload_json.asp',
		fileManagerJson : 'KEditor_keyicms/asp/file_manager_json.asp',
		allowFileManager : true
	});
	K('#image').click(function() {
		editor.loadPlugin('image', function() {
			editor.plugin.imageDialog({
				showRemote : true,
				imageUrl : K('#MapLogo').val(),
				clickFn : function(url, title, width, height, border, align) {
					K('#MapLogo').val(url);
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
	If InStr(CompanyAdminPurview, "|8,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row ">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> <%=Keyicms_Lang_Copyright(14)%></li>
                <li><%=Keyicms_Lang_Copyright(2)%></li>
                <li class="active">百度地图</li>
            </ol>	
        </div>	
        <div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-gears"></i><span class="break"></span><strong>百度地图信息设置</strong></h2>
                    </div>
                    
                    <div class="panel-body">
                    <form name="editForm" method="post" action="?Action=SaveEdit" class="form-horizontal">
					<%
                    If Conn.execute("Select * from keyicms_Map").eof Then
                        Call AKEdit()
                    Else
                        Call MapEdit()
                    End If
                    %>
                    </form>
                    </div>
                </div>		
            </div>
        </div>
    </div>
    <!-- end: Content -->
	<!--#Include file="Keyicms_End.Asp"-->				
</div>
<!--#Include file="Keyicms_js.Asp"-->

<script type="text/javascript" src="http://api.map.baidu.com/api?v=2.0&ak=<%=AK%>"></script>
<script type="text/javascript">

var point, level
var MapX, MapY, MapLevel
MapX = "<%=MapX%>"
MapY = "<%=MapY%>"
MapLevel = "<%=MapLevel%>"

var map = new BMap.Map("MapContent");
if (MapX=="" || MapY=="" || MapLevel=="")
{
	point = "成都";
	level = 12;
}
else
{
	var point = new BMap.Point(MapX, MapY);
	level = MapLevel;
	
	var marker = new BMap.Marker(point);
	map.addOverlay(marker);
	marker.disableDragging();
}
map.centerAndZoom(point,level);

map.enableScrollWheelZoom();
map.enableContinuousZoom();
map.addEventListener("zoomend", function(){document.editForm.MapLevel.value = this.getZoom()});

var textMenu = new BMap.ContextMenu();
var txtMenuItem = [
{text:'放大',callback:function(){map.zoomIn()}},
{text:'缩小',callback:function(){map.zoomOut()}},
{text:'添加标注',callback:function(p){map.clearOverlays();var marker = new BMap.Marker(p), px = map.pointToPixel(p);map.addOverlay(marker);marker.enableDragging();document.editForm.MapX.value = p.lng;document.editForm.MapY.value = p.lat}}
];
for(var i=0; i < txtMenuItem.length; i++){
	textMenu.addItem(new BMap.MenuItem(txtMenuItem[i].text,txtMenuItem[i].callback,100));
	if(i==1) {textMenu.addSeparator();}
}
map.addContextMenu(textMenu);

function searchLocation(){
	var site = $("#InputSite").val();
	var option = {
		onSearchComplete: function(results){
			if (local.getStatus() == BMAP_STATUS_SUCCESS){
				if(results.getCurrentNumPois()){
					map.centerAndZoom(results.getPoi(0).point,15);
				}
			}else{
				alert("没有检索到您输入的地址");
			}
		}
	};
	var local = new BMap.LocalSearch(map,option);
	local.search(site);
}
</script>
</body>
</html>
<%
Sub AKEdit()
	Action = request.QueryString("Action")
	If Action = "SaveEdit" Then
		If Trim(request.Form("AK"))="" Then
			Call SweetAlert("warning", "操作失败！", "请填写AK密钥！", "false", "history.back()")
			response.End
		End If
	
		Dim rs, sql
		Set rs = server.CreateObject("adodb.recordset")
		sql = "select * from keyicms_Map"
		rs.Open sql, conn, 1, 2
		rs.addnew
		rs("AK") = request.Form("AK")
		rs.update
		rs.close : set rs = nothing
		Call SweetAlert("success", "操作成功", "AK密钥设置成功！", "false", "location.replace('Ky_Map.Asp')")
	Else
	%>
        <div class="form-group">
            <label class="col-md-2 control-label" for="text-input"></label>
            <div class="col-md-6">
                <p class="form-control-static">由于你是第一次使用，请设置你的百度地图的AK密钥！</p>
            </div>
        </div>
        <div class="form-group">
            <label class="col-md-2 control-label" for="text-input">AK密钥：</label>
            <div class="col-md-4">
                <input type="text" id="AK" name="AK" class="form-control">
                <span class="help-block"></span>
            </div>
        </div>
        <div class="form-group">
            <div class="col-md-2 control-label"></div>
            <div class="col-md-10">
                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存AK密钥">
                <input type="button" class="btn btn-danger btn-lg" value="<%=Keyicms_Lang_System(32)%>" onClick="history.back(-1)">
            </div>	 
        </div>	
	<%
	End If
End Sub
%>

<%
Sub MapEdit()
	Action = request.QueryString("Action")
	If Action = "SaveEdit" Then
		If Trim(request.Form("MapX")) = "" and Trim(request.Form("MapY")) = "" and Trim(request.Form("MapLevel")) = "" Then
			Call SweetAlert("warning", "操作失败！", "坐标和级别都不能为空！", "false", "history.back()")
			response.End
		End If

		Dim rs, sql
		Set rs = server.CreateObject("adodb.recordset")
		sql = "select top 1 * from keyicms_Map"
		rs.Open sql, conn, 1, 2
		rs("MapX") = request.Form("MapX")
		rs("MapY") = request.Form("MapY")
		rs("MapLogo") = request.Form("MapLogo")
		rs("MapLevel") = request.Form("MapLevel")
		rs("AK") = request.Form("AK")
		rs.update
		rs.close : set rs = nothing
		Call SweetAlert("success", "操作成功", "百度地图信息设置成功！", "false", "location.replace('Ky_Map.Asp')")
	Else
		Set rs = server.CreateObject("adodb.recordset")
		sql = "select top 1 * from keyicms_Map"
		rs.Open sql, conn, 1, 2
		MapX = rs("MapX")
		MapY = rs("MapY")
		MapLevel = rs("MapLevel")
		MapLogo = rs("MapLogo")
		AK = rs("AK")
		%>
        <div id="map">
            <div id="MapContent" class="jvectormap-container" style="width: 100%; height: 400px;"></div>
        </div>
        <div class="form-group">
            <label class="col-md-2 control-label" for="text-input"></label>
            <div class="col-md-4">
                <div class="input-group">
                    <input type="text" id="InputSite" name="InputSite" class="form-control">
                    <span class="input-group-btn">
                        <button type="button" id="Search" class="btn btn-default" onClick="searchLocation()">搜索</button>
                    </span>
                </div>                               
                <span class="help-block">在你想要标注的地点上点击右键 - （选择）添加标注</span>
            </div>
        </div>
        <div class="form-group">
            <label class="col-md-2 control-label" for="text-input">坐标（X）：</label>
            <div class="col-md-4">
                <input type="text" id="MapX" name="MapX" class="form-control" value="<%=MapX%>" readonly>
            </div>
        </div>
        <div class="form-group">
            <label class="col-md-2 control-label" for="text-input">坐标（Y）：</label>
            <div class="col-md-4">
                <input type="text" id="MapY" name="MapY" class="form-control" value="<%=MapY%>" readonly>
            </div>
        </div>

        <div class="form-group">
            <label class="col-md-2 control-label" for="text-input">地图级别：</label>
            <div class="col-md-4">
                <input type="text" id="MapLevel" name="MapLevel" class="form-control" value="<%=MapLevel%>" readonly>
                <span class="help-block"></span>
            </div>
        </div>
        <div class="form-group">
            <label class="col-md-2 control-label" for="text-input">地图LOGO：</label>
            <div class="col-md-4">
                <div class="input-group">
                    <input type="text" id="MapLogo" name="MapLogo" class="form-control" value="<%=MapLogo%>">
                    <span class="input-group-btn">
                        <button type="button" id="image" class="btn btn-default">上传图片</button>
                    </span>
                </div>
                <span class="help-block"></span>
            </div>
        </div>
        <div class="form-group">
            <label class="col-md-2 control-label" for="text-input">AK密钥：</label>
            <div class="col-md-4">
                <input type="text" id="AK" name="AK" class="form-control" value="<%=AK%>">
            </div>
        </div>
        <div class="form-group">
            <div class="col-md-2 control-label"></div>
            <div class="col-md-10">
                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存百度地图信息">
                <input type="button" class="btn btn-danger btn-lg" value="<%=Keyicms_Lang_System(32)%>" onClick="history.back(-1)">
            </div>	 
        </div>	
	<%
	End If
End Sub
%>