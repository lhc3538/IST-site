function refreshimg(){document.all.CodeImg.src="../Include/CheckCode/CheckCode.asp"+"?r="+Math.random();}//防火狐不刷新
function GetPinyin(){
	var SortName = $('#SortName').val();
	var str = '';
	$.ajax({
		type: "get",
		url: "GetPinyin.Asp",  
		data: "SortName="+SortName,
		cache: false,
		beforeSend: function(){
			$("#Pinyin").html("请等待...");
		},
		success: function(Data){
			$("#FolderName").val(Data);
			$("#Pinyin").html("类别名称转换为拼音");
		}
	})
}

$("#Show").change(function(){
	$("[name=ShowFlag]").toggle()
})

$("#Hide").change(function(){
	$("[name=ShowFlag]").hide();
})

$("input[data-name=TabData]").change(function(){
	dataCh = $(this).attr("data-ch")
	if($("input[data-name=TabData][data-ch="+dataCh+"]:checked").attr("data-value") == 1)
	{
		$("[data-div=" + dataCh + "1]").removeAttr("style")
		$("[data-div=" + dataCh + "2]").css("display","none")
	}
	else
	{
		$("[data-div=" + dataCh + "1]").css("display","none")
		$("[data-div=" + dataCh + "2]").removeAttr("style")
	}
})

function ConfirmDel(message)
{
   if (confirm(message))
   {
      document.formDel.submit()
   }
}

function voidNum(argValue)
{
   var flag1=false;
   var compStr="1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-";
   var length2=argValue.length;
   for (var iIndex=0;iIndex<length2;iIndex++)
   {
	   var temp1=compStr.indexOf(argValue.charAt(iIndex));
	   if(temp1==-1)
	   {
	      flag1=false;
			break;
	   }
	   else
	   { flag1=true; }
   }
   return flag1;
}

function AddSort()
{
	document.editForm.SortID.value = $("[name=SortSelect]").children("option:selected").attr("value")
	document.editForm.SortPath.value = $("[name=SortSelect]").children("option:selected").attr("valuepath")
}

function CheckOther()
{
	e = $("input[name=SelectID]")
	for(var i = 0; i < e.length; i++)
	{
		if (e[i].checked==false)
		{
			e[i].checked = true;
		}
		else
		{
			e[i].checked = false;
		}
	}
}

function CheckAll()
{
	e = $("input[name=SelectID]")
	for(var i = 0; i < e.length; i++)
	{
		e[i].checked = true;
	}
}

function CheckPurviewO()
{
	e = $("input[name=Purview]")
	for(var i = 0; i < e.length; i++)
	{
		if (e[i].checked==false)
		{
			e[i].checked = true;
		}
		else
		{
			e[i].checked = false;
		}
	}
}

function CheckPurview()
{
	e = $("input[name=Purview]")
	for(var i = 0; i < e.length; i++)
	{
		e[i].checked = true;
	}
}

function ConfirmDelSort(Result,ID)
{
	swal({
		title: "提示",
		text: "是否确定删除本类、子类及下属所有信息？",
		type: "warning",
		showCancelButton: true,
		confirmButtonColor: "#DD6B55",
		confirmButtonText: '确定',
		cancelButtonText: '取消',
	},function(){
		window.location.href = Result+".asp?Action=Del&ID="+ID
	});
}

function AdminOut()
{
	swal({
		title: "提示",
		text: "是否确定退出？",
		type: "warning",
		showCancelButton: true,
		confirmButtonColor: "#DD6B55",
		confirmButtonText: '确定',
		cancelButtonText: '取消',
	},function(){
		window.location.href = "CheckAdmin.asp?AdminAction=Out"
	});
}

function test(ID, DataForm)
{
	swal({
		title: "提示",
		text: "确定删除？",
		type: "warning",
		showCancelButton: true,
		confirmButtonColor: "#DD6B55",
		confirmButtonText: '确定',
		cancelButtonText: '取消',
	},function(){
		window.location.href = "DelContent.Asp?Result="+DataForm+"&SelectID="+ID
	});
}

function test_0()
{
	var Result = $("input[name=Result]").val()
	var e = $("input[name=SelectID]")
	var SelectID = ""
	for(var i = 0; i < e.length; i++)
	{
		if(e[i].checked == true)
		{
			SelectID = e[i].value + "," + SelectID
		}
	}

	if(SelectID == ""){
		swal("提示", "你还未选择任何信息！", "warning");
		return false
	}

	swal({
		title: "提示",
		text: "确认批量删除选中的信息？",
		type: "warning",
		showCancelButton: true,
		confirmButtonColor: "#DD6B55",
		confirmButtonText: '确定',
		cancelButtonText: '取消',
		closeOnConfirm: false,
	},function(){
		window.location.href = "DelContent.Asp?Result="+Result+"&SelectID="+SelectID
	});
}

function test_1(ID)
{
	swal({
		title: "提示",
		text: "确定卸载？",
		type: "warning",
		showCancelButton: true,
		confirmButtonColor: "#DD6B55",
		confirmButtonText: '确定',
		cancelButtonText: '取消',
	},function(){
		window.location.href = "DelContent.Asp?Result=Template&SelectID="+ID
	});
}

function HtmlSite(Ke01)
{
	swal({
		title: "提示",
		text: "生成全站静态页面会消耗更多的时间，确定继续？",
		type: "warning",
		showCancelButton: true,
		confirmButtonColor: "#DD6B55",
		confirmButtonText: '确定',
		cancelButtonText: '取消',
	},function(){
		$('#html_Bar').attr('src', "Html_Site.Asp?Ke01="+Ke01);
	});
}

function num_1()
{
		var num_1=document.getElementById("Num_1").value;
		var num_1_str=document.getElementById("num_1_str");
		var str;
		str = "";
		for(var i=0;i<num_1;i++)
		{	
			str = str + "<div style='margin-bottom:10px; overflow:hidden;'>";
			str = str + "<div class='col-md-5'>";
			str = str + "	<div class='input-group'>";
			str = str + "		<span class='input-group-addon'>属性名称</span>";
			str = str + "		<input type='text' class='form-control' name='Attribute"+(parseInt(i))+"' id='Attribute"+(parseInt(i))+"' value=''>"
			str = str + "	</div>";
			str = str + "</div>";
			str = str + "<div class='col-md-5'>";
			str = str + "	<div class='input-group'>";
			str = str + "		<span class='input-group-addon'>属性值</span>";
			str = str + "		<input type='text' class='form-control' name='Attribute"+(parseInt(i))+"_value' id='Attribute"+(parseInt(i))+"_value' value=''>"
			str = str + "	</div>";
			str = str + "</div>";
			str = str + "</div>";
		}
		num_1_str.innerHTML=str;
}

function num_1_1()
{
	var num_1=document.getElementById("Num_1").value;
	var num_1_str=document.getElementById("num_1_str");
	var str;
	str = "<div style='margin-bottom:10px; overflow:hidden;'>";
	str = str + "<div class='col-md-5'>";
	str = str + "	<div class='input-group'>";
	str = str + "		<span class='input-group-addon'>属性名称</span>";
	str = str + "		<input type='text' class='form-control' name='Attribute"+(parseInt(num_1))+"' id='Attribute"+(parseInt(num_1))+"' value=''>"
	str = str + "	</div>";
	str = str + "</div>";
	str = str + "<div class='col-md-5'>";
	str = str + "	<div class='input-group'>";
	str = str + "		<span class='input-group-addon'>属性值</span>";
	str = str + "		<input type='text' class='form-control' name='Attribute"+(parseInt(num_1))+"_value' id='Attribute"+(parseInt(num_1))+"_value' value=''>"
	str = str + "	</div>";
	str = str + "</div>";
	str = str + "</div>";
	num_1_str.innerHTML=num_1_str.innerHTML+str;
	document.getElementById("Num_1").value=(parseInt(num_1)+1);
}

function GetText(num){
	var obj = $(".ke-edit-iframe").contents().find(".ke-content")
	var tempValue = obj.html().replace(/<\/?[^>]+>/g, "");
	tempValue = tempValue.replace(/[ | ]*\n/g, "");
	tempValue = tempValue.replace(/\n|\r|(\r\n)|(\u0085)|(\u2028)|(\u2029)/g, "");
	tempValue = tempValue.replace(/&nbsp;/g,"")
	$("textarea[name='Text']").val( tempValue.substring(0,num) )
}

function OnKeyDown(){
	if(event.keyCode==13) event.returnValue=false
}

function OnChange(o, Sequence){
	if(/\D/.test(o.value) || o.value == ""){
		swal({
			title: "提示",
			text: "必须为整数且不能为空！",
			type: "warning",
			showCancelButton: false,
			confirmButtonColor: "#DD6B55",
			confirmButtonText: '确定',
		});
		o.value = Sequence;
	}
}