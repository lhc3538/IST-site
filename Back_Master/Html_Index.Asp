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
            <div class="panel-body">
                <div class="progress">
                    <div style="width: 0;" aria-valuemax="100" aria-valuemin="0" aria-valuenow="60" role="progressbar" class="progress-bar" id="progress-bar">
                        0%
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>

<%Ke01=request.QueryString("Ke01")%>
<script>
$(window.parent.document).find("#html_Bar").css("height", $(".row").css("height"))
FileName = "Index,Contact,Message".split(",")

progress(1)
function progress(i)
{
	if(i > FileName.length) return;
	$.ajax({
		type: "Get",
		url: "../Html/html_ajax.Asp",
		data: "Table=Index&FileName="+FileName[i-1]+"&Ke01=<%=Ke01%>",
		cache: false,
		
		success: function(Msg){
			Percent = parseInt(i / (FileName.length) * 100) + "%";
			var oProgress = document.getElementById("progress-bar");
			$(oProgress).css("width", Percent)
			$(oProgress).html(Percent)
			progress(i + 1)
		}
	})
}
</script>