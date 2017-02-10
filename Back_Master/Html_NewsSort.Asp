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

<%
Call SiteInfo()
Ke01 = request.QueryString("Ke01")
Set rshtml = Server.CreateObject("ADODB.Recordset")
sql = "select * from keyicms_NewsSort order by ID desc"
rshtml.Open sql, conn, 1, 1
Do While Not rshtml.EOF
	ID = rshtml("ID")&","&ID
	rshtml.movenext
Loop
rshtml.Close : Set rshtml = Nothing
%>
<script>
$(window.parent.document).find("#html_Bar").css("height", $(".row").css("height"))
ID = "<%=ID%>"
IDarr = ID.split(",")
var oProgress = document.getElementById("progress-bar");
if(ID == "")
{
	Percent = "100%";
	$(oProgress).css("width", Percent).html(Percent);
}
else{
	progress(1)
}

function progress(i)
{
	if(i > IDarr.length) return;
	$.ajax({
		type: "Get",
		url: "../Html/html_ajax.Asp",  
		data: "Table=NewsSort&ID="+IDarr[i-1]+"&Ke01=<%=Ke01%>",
		cache: false,
		success: function(ErrMsg){
			Percent = parseInt(i / (IDarr.length) * 100) + "%";
			$(oProgress).css("width", Percent).html(Percent)
			progress(i + 1)
		}
	})
}
</script>