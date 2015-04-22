var bTrelloExporterLoaded=false;

function TrelloExportLoader()
{
//    console.log('Inject TrelloExport');
    $('.js-share').on('mouseup', function () {
        if(bTrelloExporterLoaded===true) return;
        bTrelloExporterLoaded=true;
        setTimeout(function(){addExportLink();}, 500);
    });      
}

chrome.extension.sendMessage({}, function(response) {
	var readyStateCheckInterval = setInterval(function() {
        
	if (document.readyState === "complete") {

		var addButtonInterval = setInterval(function(){
			
			if (!$('form').find('.trelloexport').length)
			{
				clearInterval(readyStateCheckInterval);
				clearInterval(addButtonInterval);
			}
			TrelloExportLoader();
			
		}, 5);
	
	}
	}, 10);
});

/* handle button init after page change */
var oldLocation = location.href;
setInterval(function() {
	  if(location.href != oldLocation) {

		  $('.trelloexport').remove();
          bTrelloExporterLoaded=false;
          
		  setTimeout(function() {
              // console.log('TrelloExport changed url from ' + oldLocation + ' to ' + location.href + ', btn is length ' + $('.trelloexport').length);
              if (!$('form').find('.trelloexport').length) {
                  oldLocation = location.href;
              // console.log('save old location and exit: ' + oldLocation);
                 TrelloExportLoader();
                 return;
                }
		  },1000);
	  }
}, 500);


// Add a Export Excel button to the DOM and trigger export if clicked
function addExportLink() { 
    var $js_btn = $('a.js-export-json'); // Export JSON link
    
    $('.trelloexport').remove();
    bTrelloExporterLoaded=false;
    
    if ($('form').find('.trelloexport').length) return;

    if ($js_btn.length) $excel_btn = $('<a>')
        .attr({
        class: 'trelloexport',
        href: '#',
        target: '_blank',
        title: 'TrelloExport'
    })
        .text('TrelloExport')
        .click(TrelloExportOptions)
        .insertAfter($js_btn.parent())
        .wrap(document.createElement("li"));
    
    bTrelloExporterLoaded=true;
    setTimeout(function(){bTrelloExporterLoaded=false;}, 1000);
}
