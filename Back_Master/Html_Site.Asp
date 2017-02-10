<!--#include file="../Include/Const.Asp" -->
<!--#include file="../Include/ConnSiteData.Asp" -->
<!--#include file="../Include/System_Language.asp" -->
<!--#include file="../Include/Version.Asp" -->
<!--#include file="CheckAdmin.Asp"-->

<!--#Include file="Keyicms_css.Asp"-->
<div class="row" style=" background:#eee;">	
    <div class="col-md-12">
        <div class="panel panel-default">
            <div class="panel-heading">
                <h2><i class="fa fa-tasks"></i>页面生成进度</h2>
            </div>
            <div class="panel-body" id="progress">
                <div class="progress">
                    <div aria-valuemax="100" aria-valuemin="0" aria-valuenow="60" role="progressbar" class="progress-bar" id="progress-bar">
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>

<%
Ke01 = request.QueryString("Ke01")
%>
<script>
$(window.parent.document).find("#html_Bar").css("height", $(".row").css("height"))
$.ajax({
	type: "Get",
	url: "../Html/html_To_Site.Asp",  
	data: "Ke01=<%=Ke01%>",
	cache: false, 
	beforeSend: function(){
		$("#progress-bar").css("background","url(ke01/img/progress.gif)")
		$("#progress-bar").css("width","100%")
		$("#progress-bar").html("生成中...")
	},
	success: function(Msg){
		$("#progress-bar").css("background","")
		$("#progress-bar").html("全站静态生成成功")
		$("#progress").append("<div class='form-group'><div class='col-md-12'><p class='form-control-static'>"+Msg+"</p></div></div>")
		
		$(window.parent.document).find("#html_Bar").css("height", $(".row").css("height"))
	}
})
</script>