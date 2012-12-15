// Variables
var reg = /[\(](\x3f|\d*\.?\d+)([\)])\s?/m;



//what to do when DOM loads
$(function(){
    $('.js-share').live('mouseup',function(){
		setTimeout(checkExport);
	});
});

//for export
var $excel_btn,$excel_dl;
window.URL = window.webkitURL || window.URL;

function checkExport() {
	if($('form').find('.js-export-excel').length) return;
	var $js_btn = $('a.js-export-json');
	console.log($js_btn.length);
	
	if($js_btn.length)
		$excel_btn = $('<a>')
			.attr({
				class: 'js-export-excel',
				href: '#',
				target: '_blank',
				title: 'Open downloaded file with Excel'
			})
			.text('Export Excel')
			.click(showExcelExport)
			.insertAfter($js_btn)
			.wrap(document.createElement("li"));
}

function showExcelExport() {
	$excel_btn.text('Generating...');

	$.getJSON($('a.js-export-json').attr('href'), function(data) {
		var s = '<table id="export" border=1>';
		s += '<tr><th>Points</th><th>Story</th><th>Description</th></tr>';
		$.each(data['lists'], function(key, list) {
			var list_id = list["id"];
			s += '<tr><th colspan="3">' + list['name'] + '</th></tr>';

			$.each(data["cards"], function(key, card) {
				if (card["idList"] == list_id) {
					var title = card["name"];
					var parsed = title.match(reg);
					var points = parsed?parsed[1]:'';
					title = title.replace(reg,'');
					s += '<tr><td>'+ points + '</td><td>' + title + '</td><td>' + card["desc"] + '</td></tr>';
				}
			});
			s += '<tr><td colspan=3></td></tr>';
		});
		s += '</table>';

		var blob = new Blob([s], { type: 'application/ms-excel' });

		var board_title_reg = /.*\/board\/(.*)\//;
		var board_title_parsed = document.location.href.match(board_title_reg);
		var board_title = board_title_parsed[1];

		$excel_btn
			.text('Excel')
			.after(
				$excel_dl=$('<a>')
					.attr({
						download: board_title + '.xls',
						href: window.URL.createObjectURL(blob)
					})
			);

		var evt = document.createEvent('MouseEvents');
		evt.initMouseEvent('click', true, true, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
		$excel_dl[0].dispatchEvent(evt);
		$excel_dl.remove()

	});

	return false
};
