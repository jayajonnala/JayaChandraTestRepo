$(document).ready(function(){
	// Tooltip only Text
	$('.masterTooltip').hover(function(){
		// Hover over code
		var title = $(this).attr('title');
		$(this).data('tipText', title).removeAttr('title');
		$('<p class="tooltip"></p>')
		.text(title)
		.appendTo('body')
		.fadeIn('slow');
	}, function() {
		// Hover out code
		$(this).attr('title', $(this).data('tipText'));
		$('.tooltip').remove();
	}).mousemove(function(e) {
		var mousex = e.pageX + 20; //Get X coordinates
		var mousey = e.pageY + 10; //Get Y coordinates
		$('.tooltip')
		.css({ top: mousey, left: mousex })
	});
	
	$('ul.iter-sub-head li').click(function(){
		var liElementId = $(this).attr('id');
		var h3ElementId = '#'+liElementId.split("-")[2];
		var focusId = '#'+h3ElementId.split("Iteration")[1]
		$('html, body').animate({
			scrollTop: $(focusId).offset().top
		}, 2000);
		$(h3ElementId).trigger("evt_accordion")
		return false
	});
	
	$('ul.iter-sub-head').find('li.fail a').css({'font-weight':'bold','color':'#fe0400'});	
    // Iteration Selecetion - Start	
		$('table.stat-status td.fail').bind('click', failTitle);
        $('table.stat-status td.pass').bind('click', passTitle);
        $('table.stat-status td.total').bind('click', totalTitle);
		function failTitle() { allnormalTitle(); $('ul.iter-sub-head').find('li.fail a').css({'font-weight':'bold','color':'#fe0400'});};
        function passTitle() {allnormalTitle(); $('ul.iter-sub-head').find('li.pass a').css({'font-weight':'bold','color':'#18722a'});};
        function totalTitle() {allnormalTitle(); $('ul.iter-sub-head').find('li.total a').css({'font-weight':'bold','color':'#111111'});};
		function allnormalTitle() {$('ul.iter-sub-head').find('li a').css({'font-weight':'normal','color':'#15527f'});};
	// Iteration Selecetion - End
    
    $('ul.case-sub-head').find('li.fail a').css({'font-weight':'bold','color':'#fe0400'});	
    // Casestudy Selecetion - Start	
	$('table.case-status td.fail').click(function(){
		var tdElementId = $(this).attr('id');
		var focusId = '#'+tdElementId.split("-")[1];
		var ulElementId = "TestCaseList-" + tdElementId.split("-")[1];
		$('ul.case-sub-head#'+ulElementId).find('li a').css({'font-weight':'normal','color':'#15527f'});
		$('ul.case-sub-head#'+ulElementId).find('li.fail a').css({'font-weight':'bold','color':'#fe0400'});
		$(window).scrollTop($(focusId).offset().top)
	});
	$('table.case-status td.pass').click(function(){
		var tdElementId = $(this).attr('id');
		var focusId = '#'+tdElementId.split("-")[1];
		var ulElementId = "TestCaseList-" + tdElementId.split("-")[1];
		$('ul.case-sub-head#'+ulElementId).find('li a').css({'font-weight':'normal','color':'#15527f'});
		$('ul.case-sub-head#'+ulElementId).find('li.pass a').css({'font-weight':'bold','color':'#18722a'});
		$(window).scrollTop($(focusId).offset().top)
	});
	$('table.case-status td.total').click(function(){
		var tdElementId = $(this).attr('id');
		var focusId = '#'+tdElementId.split("-")[1];
		var ulElementId = "TestCaseList-" + tdElementId.split("-")[1];
		$('ul.case-sub-head#'+ulElementId).find('li a').css({'font-weight':'normal','color':'#15527f'});
		$('ul.case-sub-head#'+ulElementId).find('li.total a').css({'font-weight':'bold','color':'#111111'});
		$(window).scrollTop($(focusId).offset().top)
	});
	// Casestudy Selecetion - End
	$('#setting-trigger').click(function(){
		$(this).next('#setting-content').toggle();
		$(this).toggleClass('active');          
		if ($(this).hasClass('active')) $(this).find('span').html('&#x25B2;')
		else $(this).find('span').html('&#x25BC;')
    })
	
	$('#setting-content').focusout(function(){
		$(this).toggle();
		$('#setting-trigger').toggleClass('active');          
		if ($('#setting-trigger').hasClass('active')) $('#setting-trigger').find('span').html('&#x25B2;')
		else $('#setting-trigger').find('span').html('&#x25BC;')
    })
	
	$('#bar').click(function(){
		$('#piehead').css({'display':'none'});
		$('#piegraph').css({'display':'none'});
		$('#bargraph').css({'display':'inline'});
		for (var j = 0; j < $(".vert-bars .data li").length; ++j) {
		   $(".vert-bars .bar:eq(" + j + ")").animate({
			   height: $(".vert-bars .value:eq(" + j + ")").attr('value')
			}, 500);
		}
	});
	
	$('#pie').click(function(){
		$('#piehead').css({'display':'inline'});
		$('#bargraph').css({'display':'none'});
		$('#piegraph').css({'display':'inline'});
	});
	
	$('#srd').click(function(){
		if($('#srd').prop("checked")){
			$('#srate').css({'display':'table-cell'});
		}else{
			$('#srate').css({'display':'none'});
		}
	});
	$('#frd').click(function(){
		if($('#frd').prop("checked")){
			$('#frate').css({'display':'table-cell'});
		}else{
			$('#frate').css({'display':'none'});
		}
	});
});	