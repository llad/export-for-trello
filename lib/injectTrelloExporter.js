var bTrelloExporterLoaded=false;

function TrelloExporterLoader()
{
//    console.log('Inject TrelloExport');
    $('.js-share').on('mouseup', function () {
    if(bTrelloExporterLoaded==true) return;
    bTrelloExporterLoaded=true;
    //console.log('TrelloExport 1.9.2 DOM bTrelloExporterLoaded');
    setTimeout(function(){addExportLink();}, 500);
});  
    
}

chrome.extension.sendMessage({}, function(response) {
	var readyStateCheckInterval = setInterval(function() {
        
//		console.log('document.readyState=' + document.readyState);
        
//        console.log('bTrelloExporterLoaded: ' + bTrelloExporterLoaded);
        
	if (document.readyState === "complete") {
		//clearInterval(readyStateCheckInterval);

		var addButtonInterval = setInterval(function(){
			
			if (!$('form').find('.js-export-excel').length)
			{
				clearInterval(readyStateCheckInterval);
				clearInterval(addButtonInterval);
			}
			TrelloExporterLoader();
			
		}, 5);
	
	}
	}, 10);
});

/* handle button init after page change */
var oldLocation = location.href;
setInterval(function() {
	  if(location.href != oldLocation) {

		  $('.js-export-excel').remove();
          bTrelloExporterLoaded=false;
          
		  setTimeout(function(){

//		  console.log('TrelloExport changed url from ' + oldLocation + ' to ' + location.href + ', btn is length ' + $('.js-export-excel').length);
		 
         if (!$('form').find('.js-export-excel').length) {
              oldLocation = location.href;
//             console.log('save old location and exit: ' + oldLocation);
             TrelloExporterLoader();
             return;
         }
              
		  },1000);


	  }
}, 500);


// Add a Export Excel button to the DOM and trigger export if clicked
function addExportLink() { 
    var $js_btn = $('a.js-export-json'); // Export JSON link
    
    $('.js-export-excel').remove();
    bTrelloExporterLoaded=false;
    
    // See if our Export Excel is already there
    if ($('form').find('.js-export-excel').length) return;

//    console.log('addExportLink');
    
    // The new link/button
    if ($js_btn.length) $excel_btn = $('<a>')
        .attr({
        class: 'js-export-excel',
        href: '#',
        target: '_blank',
        title: 'Open downbTrelloExporterLoaded file with Excel'
    })
        .text('Export Excel')
        .click(createExcelExport)
        .insertAfter($js_btn.parent())
        .wrap(document.createElement("li"));
    
    bTrelloExporterLoaded=true;
    setTimeout(function(){bTrelloExporterLoaded=false;}, 1000);
}
