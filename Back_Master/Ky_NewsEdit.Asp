<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/system_language.asp" -->
<!--#include file="CheckAdmin.Asp"-->
<!--#include file="../Html/html_Config.Asp"-->
<!--#include file="../Html/html_Function.Asp"-->
<!--#include file="../Html/html_To_News.Asp"-->
<%
Dim Result
Result = request.QueryString("Result")
Dim ID, NewsName, ViewFlag, MobileFlag, NoticeFlag, PicFlag, ClassSeo, SortName, SortID, SortPath, Text, Bold, Color, SmallPic, ClickNumber,AddTime
Dim Origin, Content, SeoKeywords, SeoDescription
ID = request.QueryString("ID")
Call NewsEdit()
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
	var editor = K.create('textarea[name="Content"]', {
		cssPath : 'KEditor_keyicms/plugins/code/prettify.css',
		uploadJson : 'KEditor_keyicms/asp/upload_json.asp',
		fileManagerJson : 'KEditor_keyicms/asp/file_manager_json.asp',
		allowFileManager : true,
		afterBlur : function() {
		this.sync();
		K.ctrl(document, 13, function() {
		K('form[name=editForm]')[0].submit();
		});
		K.ctrl(this.edit.doc, 13, function() {
		K('form[name=editForm]')[0].submit();
		});
		}
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
	If InStr(CompanyAdminPurview, "|2,") = 0 Then
		response.Write ("<br /><br /><div style='height:100%; text-align:center' class='main'><font style=""color:red; font-size:14px;"")>"&Keyicms_Lang_System(26)&"</font></div>")
		response.End
	End If
	%>
    <div class="main">
        <div class="row ">
            <ol class="breadcrumb">
                <li><i class="fa fa-home"></i> <%=Keyicms_Lang_Copyright(14)%></li>
                <li><%=Keyicms_Lang_Copyright(2)%></li>
                <li class="active"><a href="">功能模块</a></li>
                <li class="active">信息管理</li>
            </ol>	
        </div>	
        <div class="row ">
            <div class="col-lg-12 hideInIE8">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-edit"></i><span class="break"></span><strong>信息管理</strong></h2> 
                    </div>
                    <div class="panel-body">
                        <a class="ShortCut1" href="Ky_NewsEdit.Asp?Result=Add">  
                            <i class="fa fa-indent"></i>
                            <div class="title">添加信息</div>
                        </a>					
                        <a class="ShortCut2" href="Ky_News.Asp">  
                            <i class="fa fa-list"></i>
                            <div class="title">信息列表</div>
                        </a>	
                        <a class="ShortCut3" href="Ky_NewsSort.Asp">  
                            <i class="fa fa-tasks"></i>
                            <div class="title">信息分类</div>
                        </a>													
                    </div>
                </div>
            </div>
        </div>
        <div class="row">
            <div class="col-md-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <h2><i class="fa fa-gears"></i><span class="break"></span><strong><%If Result = "Add" then%>添加<%ElseIf Result = "Modify" then%>修改<%End If%>信息</strong></h2>
                    </div>
                    <div class="panel-body">
                    <form name="editForm" method="post" action="?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" class="form-horizontal">
                        <div class="form-group">
                            <label class="col-md-2 control-label">信息名称：</label>
                            <div class="col-md-4">
                                <input name="NewsName" type="text" id="NewsName" class="form-control" value="<%=NewsName%>" maxlength="100">
                            </div>
                        </div> 
                        
                        <%If EnableSeo = True Then%>
                        <div class="form-group">
                            <label class="col-md-2 control-label">关键词：</label>
                            <div class="col-md-6">
                            	<input name="SeoKeywords" type="text" id="SeoKeywords" class="form-control" value="<%=SeoKeywords%>" maxlength="250">
                                <span class="help-block">SEO：多个关键词请用“,”半角逗号分开</span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label">描述：</label>
                            <div class="col-md-6">
                            	<input name="SeoDescription" type="text" id="SeoDescription" class="form-control" value="<%=SeoDescription%>" maxlength="250">
                                <span class="help-block">SEO：请填写描述</span>
                            </div>
                        </div>
                        <%End If%>
                        <div class="form-group">
                            <label class="col-md-2 control-label">标题加粗：</label>
                            <div class="col-md-10">
                                <label class="radio-inline" for="Bold1">
                                    <input type="radio" name="Bold" id="Bold1" value="Bold" <%If Bold="Bold" then response.write("checked='checked'")%>/> 是
                                </label>
                                <label class="radio-inline" for="Bold2">
                                    <input type="radio" name="Bold" id="Bold2" value="Normal" <%If Bold="Normal" or Bold<>"Bold" then response.write("checked='checked'")%> />否
                                </label>
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label">标题颜色：</label>
                            <div class="col-md-4 col-xs-11">
                                <div class="input-group">
                                    <input id="Color" name="Color" type="text" value="<%=Color%>" readonly class="form-control" />
                                    <span class="input-group-btn">
                                    	<img id="coloImg" src="ke01/img/color.png" style="width:24px; height:24px; cursor:pointer; vertical-align:middle;" onClick="Colorcd('editForm','Color','Color',event);">
                                    </span>
                                </div>
                                <span class="help-block">这里可以选择设置标题颜色</span>
                            </div>
                        </div>
                       
                        <div class="form-group">
                            <label class="col-md-2 control-label">信息类别：</label>
                            <div class="col-md-5">
                                <select size="1" class="form-control" name="SortSelect" id="SortSelect" onChange="AddSort()">
                                <option value="" valuepath="">请选择...</option>
                                <%Call NewsSort(0, 0)%>
                                </select>
                                <span class="help-block">
                                    <input name="SortID" type="hidden" id="SortID" value="<%=SortID%>" readonly class="form-control">
                                    <input name="SortPath" type="hidden" id="SortPath" value="<%=SortPath%>" readonly class="form-control">
                                </span>
                            </div>
                        </div>  
                        <div class="form-group">
                            <label class="col-md-2 control-label">信息来源：</label>
                            <div class="col-md-5">  
                                <input name="Origin" type="text"  value="<%=Origin%>" class="form-control" maxlength="100">
                                <a href="javascript:" onClick="document.editForm.Origin.value='网络转载'">网络转载</a> &nbsp; | &nbsp; 
                                <a href="javascript:" onClick="document.editForm.Origin.value='原创'">原创</a> 
                                <span class="help-block"></span>
                            </div>
                        </div>    
                        <div class="form-group">
                            <label class="col-md-2 control-label">信息内容：</label>
                            <div class="col-md-10">
                                <textarea name="Content" style="width:95%;height:400px;visibility:hidden;" class="form-control"><%=Server.HTMLEncode(Content)%></textarea>
                                <span class="help-block"></span>
                            </div>
                        </div>
                        <div class="form-group">
                            <label class="col-md-2 control-label">信息摘要：</label>
                            <div class="col-md-5">            
                                <textarea name="Text" rows="5" style="width:100%; height:60px;" class="form-control"><%=Text%></textarea>
                                <button type="button" class="btn btn-default" onClick="GetText(50)">获取内容前50字符</button>                              
                                <span class="help-block"></span>
                            </div>
                        </div>   
                        <%If EnableViews = True Then%>  
                        <div class="form-group">
                            <label class="col-md-2 control-label">信息浏览数：</label>
                            <div class="col-md-4">
                                <input type="text" name="ClickNumber" id="ClickNumber" class="form-control" value="<%=ClickNumber%>" onKeyDown="OnKeyDown()" onChange="OnChange(this, <%=ClickNumber%>)" />
                                <span class="help-block">可以手动输入浏览数量。</span>
                            </div>
                        </div> 
						<div class="form-group">
                            <label class="col-md-2 control-label">发布日期：</label>
                            <div class="col-md-4">
							<%
							if Result = "Add" then
								AddTime=now()
							end if
							%>
                                <input type="text" name="AddTime" id="AddTime" class="form-control" value="<%=AddTime%>" onKeyDown="OnKeyDown()" onChange="OnChange(this, <%=AddTime%>)" />
								<button href="javascript:" type="button" class="btn btn-default" onClick="document.editForm.AddTime.value='<%=now()%>'">当前时间</button> 
								<span class="help-block">可以手动输入发布日期。</span> 
                            </div>
                        </div> 
                        <%End If%>
                        <div class="form-group">
                            <div class="col-md-2 control-label"></div>
                            <div class="col-md-10">
                                <input name="submitSaveEdit" type="submit" id="submitSaveEdit" class="btn btn-danger btn-lg" value="保存信息">
                                <input type="button" class="btn btn-danger btn-lg" value="<%=Keyicms_Lang_System(32)%>" onClick="history.back(-1)">
                            </div>	
                        </div>
                    </form>
                    </div>
                </div>
            </div>		
        </div>
    </div>
	<!--#Include file="Keyicms_End.Asp"-->				
</div>
<!--#Include file="Keyicms_js.Asp"-->
<script>
$("[name=PicFlag]").change(function(){
	$("[name=PicDiv]").toggle()
})
</script>
<div id="Colorpanel" style="position: absolute; left: 0; top: 0; z-index: 1000;"></div>
<script>
var ColorHex=new Array('00','33','66','99','CC','FF')
var SpColorHex=new Array('FF0000','00FF00','0000FF','FFFF00','00FFFF','FF00FF')
var current=null
function intoColor(dddd,ssss,ffff)
{	
    var ColorTable=''
    for (i=0;i<2;i++)
    {
        for (j=0;j<6;j++)
        {
            ColorTable=ColorTable+'  <tr>'
            ColorTable=ColorTable+'    <td style="width:12px; height:12px; background-Color: #000000;">'
            if (i==0){
                ColorTable=ColorTable+'    <td style="width:12px; height:12px; background-Color: #'+ColorHex[j]+ColorHex[j]+ColorHex[j]+';">'
            }
            else{
                ColorTable=ColorTable+'    <td style="width:12px; height:12px; background-Color: #'+SpColorHex[j]+';">'
            }
            ColorTable=ColorTable+'    <td style="width:12px; height:12px; background-Color: #000000;">'
            for (k=0;k<3;k++)
            {
                for (l=0;l<6;l++)
                {
                    ColorTable=ColorTable+'    <td style="width:12px; height:12px; background-Color: #'+ColorHex[k+i*3]+ColorHex[l]+ColorHex[j]+';">'
                }
            }
        }
    }
    ColorTable='<table width="253" border="0" cellspacing="0" cellpadding="0" style="border: 1px #000000 solid; border-bottom: none; border-collapse: collapse;" borderColor="000000">'
    +'  <tr height="30">'
    +'    <td colspan="21" bgColor="#cccccc"><table width="100%" cellpadding="0" cellspacing="1" border="0" style="border-collapse: collapse;">'
    +'        <tr>'
    +'          <td width="3"></td>'
    +'          <td><input type="text" id="DisColor" name="DisColor" size="6" disabled style="border: solid 1px #000000; background-Color: #ffff00;"></td>'
    +'          <td width="3"></td>'
    +'          <td><input type="text" id="HexColor" name="HexColor" size="7" style="border: inset 1px; font-family: Arial;" value="#000000">&nbsp;&nbsp;&nbsp;&nbsp;系统选色版</td>'
    +'          <td width="20" style="cursor:pointer" onclick="document.getElementById(\'Colorpanel\').style.visibility=\'hidden\'" title=\' 关闭\'>╳</td>'
    +'        </tr>'
    +'      </table></td>'
    +'</table>'
    +'<table border="1" cellspacing="0" cellpadding="0" style="border-collapse: collapse; cursor: pointer;" borderColor="000000" onmouseover="doOver(event)" onmouseout="doOut()" onclick="doclick(\''+dddd+'\',\''+ssss+'\',\''+ffff+'\',event)">'
    +ColorTable+'</table>';
    Colorpanel.innerHTML=ColorTable
}
function doOver(e) {
	if( !e ) e = window.event;
	var Colorxs = e.target||e.srcElement; 
	
    if ((Colorxs.tagName=="TD") && (current!=Colorxs)) {
        if (current!=null){
            current.style.backgroundColor = current._BackGround
        }
        Colorxs._BackGround = Colorxs.style.backgroundColor
		
		document.getElementById("DisColor").style.backgroundColor= Colorxs.style.backgroundColor
		document.getElementById("HexColor").value= Colorxs.style.backgroundColor
		Colorxs.style.backgroundColor = "white"
        current = Colorxs
    }
}
function doOut() {
    if (current!=null) current.style.backgroundColor = current._BackGround
}

function doclick(dddd,ssss,ffff,e){
	if( !e ) e = window.event;
	var Colorxs = e.target||e.srcElement; 

    if (Colorxs.tagName=="TD"){
        document.getElementById(ssss).value=Colorxs._BackGround
		document.getElementById(ffff).style.color=Colorxs._BackGround
		
		Colorxs.style.backgroundColor=Colorxs._BackGround
		document.getElementById("Colorpanel").style.visibility="hidden"
        return Colorxs._BackGround
    }
}
var Colorxs

function Colorcd(dddd,ssss,ffff,e){
	var Colorpanel=document.getElementById("Colorpanel")
	
	if( !e ) e = window.event;
	var Colorxs = e.target||e.srcElement; 

	var x=e.clientX
    var y=e.clientY
	
    var rightedge = document.body.clientWidth-x;
    var bottomedge = document.body.clientHeight-y;
	if (rightedge < Colorpanel.offsetWidth){
    	var palLeft = document.body.scrollLeft + x - Colorpanel.offsetWidth;
		Colorpanel.style.left = palLeft+'px'
	}
    else{
    	var palLeft = document.body.scrollLeft + x;
		Colorpanel.style.left = palLeft+'px'
	}
    
	if (bottomedge < Colorpanel.offsetHeight){
    	var palTop = document.body.scrollTop + y - Colorpanel.offsetHeight;
		Colorpanel.style.top = palTop + 'px'
	}
    else{
		var ScrollTop = window.pageYOffset || document.documentElement.scrollTop || document.body.scrollTop || 0;   
		var palTop = ScrollTop + y;
		Colorpanel.style.top = palTop + 'px'
	}
	
	Colorpanel.style.visibility = "visible";
	
    e.cancelBubble=true
    intoColor(dddd,ssss,ffff)
    return false
}
function clear_color() {
	$('#Title').css('color','');
	$('#TitleFontColor').val('');
	$('#colorpanel_title').html(' ');
}
</script>

</body>
</html>

<%
Sub NewsEdit()
dim Action, rs, sql
Action = Request.QueryString("Action")
if Action = "SaveEdit" then
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("NewsName"))) < 3 then
		Call SweetAlert("warning", "友情提示", "信息名称不得少于3个字符！", "false", "history.back(-1)")
		response.end
    end if
	if Request.Form("SortID") = "" And Request.Form("SortPath") = "" Then
		Call SweetAlert("warning", "友情提示", "请选择所属分类！", "false", "history.back(-1)")
		response.End
	end if
	If Trim(request.Form("Content")) = "" Then
		Call SweetAlert("warning", "友情提示", "请填写信息内容！", "false", "history.back(-1)")
		response.End
	End If

    if Result = "Add" then
		sql = "select * from keyicms_News"
		rs.open sql, conn, 1, 3
		rs.addnew
		rs("NewsName") = trim(Request.Form("NewsName"))

		rs("ViewFlag") = 1	'默认显示

		if Request.Form("NoticeFlag") = 1 then
			rs("NoticeFlag") = Request.Form("NoticeFlag")
		else
			rs("NoticeFlag") = 0
		end if
		if Request.Form("MobileFlag") = 1 then
			rs("MobileFlag") = Request.Form("MobileFlag")
		else
			rs("MobileFlag") = 0
		end if
		If Request.Form("ClassSeo") = "" then
			If Request.Form("oAutopinyin") = "1" then rs("ClassSeo") = TimeToName()
			If Request.Form("oAutopinyin") = "2" then rs("ClassSeo") = Left(Pinyin(trim(request.form("NewsName"))),200)
		Else
			rs("ClassSeo") = trim(Request.form("ClassSeo"))
		End If
		rs("SortID") = Request.Form("SortID")
		rs("SortPath") = Request.Form("SortPath")
		rs("Origin") = trim(Request.Form("Origin"))
		if Request.Form("PicFlag") = 1 then
			rs("PicFlag") = Request.Form("PicFlag")
			rs("SmallPic") = Request.Form("SmallPic")
		else
			rs("PicFlag") = 0
			rs("SmallPic") = ""
		end if
		rs("Content") = trim(Request.Form("Content"))
		rs("SeoKeywords") = trim(Request.Form("SeoKeywords"))
		rs("SeoDescription") = trim(Request.Form("SeoDescription"))
		rs("Text") = trim(Request.Form("Text"))
		'rs("AddTime") = now()
		rs("UpdateTime") = now()
		rs("Bold") = trim(request.Form("Bold"))
		rs("Color") = trim(Request.Form("Color"))
		rs("AddTime") = trim(Request.Form("AddTime"))	'从客户机获取信息
		If EnableViews Then rs("ClickNumber") = trim(Request.Form("ClickNumber"))

		rs.update
		rs.close : set rs = nothing
	
		set rs = server.createobject("adodb.recordset")
		sql = "select top 1 ID,ClassSeo from keyicms_News order by ID desc"
		rs.open sql,conn,1,1
		ID = rs("ID")
		rs.close : set rs = nothing

		'Call html_To_News(ID, "PC")
		'If mStatus Then Call html_To_News(ID, "Mobile")
	end if
	if Result="Modify" then
		sql = "select * from keyicms_News where ID="&ID
		rs.open sql, conn, 1, 3
		rs("NewsName") = trim(Request.Form("NewsName"))

		rs("ViewFlag") = 1

		if Request.Form("NoticeFlag") = 1 then
			rs("NoticeFlag") = Request.Form("NoticeFlag")
		else
			rs("NoticeFlag") = 0
		end if
		if Request.Form("MobileFlag") = 1 then
			rs("MobileFlag") = Request.Form("MobileFlag")
		else
			rs("MobileFlag") = 0
		end if

	    rs("SortID") = Request.Form("SortID")
		rs("SortPath") = Request.Form("SortPath")
		rs("Origin") = trim(Request.Form("Origin"))
		if Request.Form("PicFlag") = 1 then
			rs("PicFlag") = Request.Form("PicFlag")
			rs("SmallPic") = Request.Form("SmallPic")
		else
			rs("PicFlag") = 0
			rs("SmallPic") = ""
		end if
		rs("Text") = trim(Request.Form("Text"))
		rs("Content") = trim(Request.Form("Content"))
		rs("SeoKeywords") = trim(Request.Form("SeoKeywords"))
		rs("SeoDescription") = trim(Request.Form("SeoDescription"))
		rs("Bold") = trim(request.Form("Bold"))
		rs("Color") = trim(Request.Form("Color"))
		rs("UpdateTime") = now()
		rs("AddTime") = trim(Request.Form("AddTime"))	'request从客户机获取信息
		
		If EnableViews Then rs("ClickNumber") = trim(Request.Form("ClickNumber"))
		ID = rs("ID")
		rs.update
		rs.close : set rs = nothing
		
		'Call html_To_News(ID, "PC")
		'If mStatus Then Call html_To_News(ID, "Mobile")
	end if
	Call SweetAlert("success", "操作成功", "设置成功，相关内容已更新！", "false", "location.replace('Ky_News.Asp')")
else
	if Result="Modify" then
		set rs = server.createobject("adodb.recordset")
		sql = "select * from keyicms_News where ID="&ID
		rs.open sql, conn, 1, 1
		if rs.bof and rs.eof then
			response.write ("<center>数据库记录读取错误！</center>")
			response.end
		end if
		NewsName = rs("NewsName")
		ViewFlag = rs("ViewFlag")
		ClassSeo = rs("ClassSeo")&Separated&rs("ID")
		SortName = SortText(rs("SortID"))
		SortID = rs("SortID")
		SortPath = rs("SortPath")
		Origin = rs("Origin")
		NoticeFlag = rs("NoticeFlag")
		MobileFlag = rs("MobileFlag")
		PicFlag = rs("PicFlag")
		Content = rs("Content")
		SeoKeywords = rs("SeoKeywords")
		SeoDescription = rs("SeoDescription")
		Text = rs("Text")
		Bold = rs("Bold")
		Color = rs("Color")
		SmallPic = rs("SmallPic")
		ClickNumber = rs("ClickNumber")
		AddTime = rs("AddTime")
		rs.close : set rs = nothing
	else
		Color = "rgb(0, 0, 0)"
		ClickNumber = 10
    end if
end if
End Sub

Function SortText(ID)
	Dim rs,sql
	set rs = Server.CreateObject("adodb.recordset")
	sql = "Select * From keyicms_NewsSort where ID="&ID
	rs.open sql,conn,1,1
	SortText = rs("SortName")
	rs.close : set rs = nothing
End Function

Function NewsSort(ID, ClassNum)
	Dim rs, sql
	set rs = Server.CreateObject("adodb.recordset")
	rs.open "select * from Keyicms_NewsSort where ParentID="&ID&" Order by ID asc",conn,1,2
	Do Until rs.eof
		If SortID = rs("ID") Then
			response.Write "<option value='"&rs("ID")&"' valuepath='"&rs("SortPath")&"' selected>"
		Else
			response.Write "<option value='"&rs("ID")&"' valuepath='"&rs("SortPath")&"'>"
		End If
		For ClassN = 1 to ClassNum
			response.Write "│"
		Next
		response.Write "│├ "&rs("SortName")
		response.Write "</option>"
		ClassNum = ClassNum + 1
		Call NewsSort(rs("ID"), ClassNum)
		ClassNum = ClassNum - 1
		rs.movenext
	Loop
	rs.close : set rs = nothing
End Function
%>