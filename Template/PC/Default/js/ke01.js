$(function () { 
	var page = 1;//初始化page变量 
	var i = 5;//每版放两组图片 
	var $pictureShow = $(".CpBox"); 
	var downwidth = $pictureShow.width();//获取框架内容的宽度，既每次移动的大小 
	var len = $(".CpBox ul").find('li').length;//找到一共有几组图片 
	var page_number = Math.ceil(len / i); 
	//找到一共有多少个版面 

	$(".Next").click(function () { 
		if (!$(".CpBox ul").is(":animated")) {//判断是否正在执行动画效果 
			if (page == page_number) {//已经到最后一个版面了,如果再向后，必须跳转到第一个版面。 
				$(".CpBox ul").animate({ left: '0px' }, "slow");//通过改变left值，跳转到第一个版面 
				page = 1; 
			}
			else { 
				$(".CpBox ul").animate({ left: '-=' + downwidth }, "slow");//通过改变left值，达到每次换一个版面 
				page++; 
			} 
		} 
		return false; 
	}); 
	
	$(".Prev").click(function () { 
		if (!$(".CpBox ul").is(":animated")) { 
			if (page == 1) {//已经到第一个版面了,如果再向前，必须跳转到最后一个版面 
				$(".CpBox ul").animate({ left: '-=' + downwidth * (page_number - 1) }, "slow");//通过改变left值，跳转到最后一个版面 
				page = page_number; 
			}
			else { 
				$(".CpBox ul").animate({ left: '+=' + downwidth }, "slow");//通过改变left值，达到每次换一个版面 
				page--; 
			} 
		} 
		return false; 
	}); 
	
	function refreshimg(){document.all.CodeImg.src="/Include/CheckCode/CheckCode.Asp"+"?r="+Math.random();}//防火狐不刷新
});