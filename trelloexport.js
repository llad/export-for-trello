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
			.insertAfter($js_btn.parent())
			.wrap(document.createElement("li"));
}

function showExcelExport() {
    
    
    // RegEx to find the points for users of TrelloScrum
    var pointReg = /[\(](\x3f|\d*\.?\d+)([\)])\s?/m;

	$.getJSON($('a.js-export-json').attr('href'), function(data) {
		var s = '<table id="export" border=1>';
		s += '<tr><th>Story</th><th>Description</th><th>Points</th></tr>';
		$.each(data['lists'], function(key, list) {
			var list_id = list["id"];
			s += '<tr><th colspan="3">' + list['name'] + '</th></tr>';
			    
    		$.each(data["cards"], function(key, card) {
				if (card["idList"] == list_id) {
					var title = card["name"];
					var parsed = title.match(pointReg);
					var points = parsed?parsed[1]:'';
					title = title.replace(pointReg,'');
					s += '<tr><td>' + title + '</td><td>' + card["desc"] + '</td><td>'+ points + '</td></tr>';
				}
			});
			s += '<tr><td colspan=3></td></tr>';
		});
		s += '</table>';


		var blob = new Blob([s], { type: 'application/vnd.ms-excel' });
		var board_title = data.name;
		saveAs(blob, board_title + '.xls');
        $("a.close-btn")[0].click();
    
    
    });

	return false
	
};
