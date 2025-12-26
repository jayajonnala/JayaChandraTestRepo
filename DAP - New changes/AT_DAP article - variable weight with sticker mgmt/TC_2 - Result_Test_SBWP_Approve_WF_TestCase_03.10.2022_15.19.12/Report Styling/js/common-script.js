$(document).ready(function(){
	var amountScrolled = 100;
	$(window).scroll(function() {
		if ( $(window).scrollTop() > amountScrolled ) {
			$('a.backtop').fadeIn('slow');
		} else {
			$('a.backtop').fadeOut('slow');
		}
	}); 
	$('a.backtop').hide(); 
	
	$(window).scroll(function() {
		if ( $(window).scrollTop() > amountScrolled ) {
			$('a.collapse-all').fadeIn('slow');
		} else {
			$('a.collapse-all').fadeOut('slow');
		}
	}); 

	$('a.collapse-all').hide(); 
	$('#collapseAll').click(function () { 
		$('.action-collepse.toggle_open').each(function(){$(this).trigger("evt_accordion");});
		return false;
	});
	$('#expandAll').click(function () { 
		$('.action-collepse').each(function(){if (!$(this).hasClass('toggle_open')) $(this).trigger("evt_accordion");});
		return false;
	});
	
	$('table.detail-report tfoot tr').each(function(){if ($(this).hasClass('fail')) { 
			var idtr = $(this).attr('id');
			var idh3 = '#' + idtr.split("End")[0];
			$(idh3).addClass('fail');
	}});
});