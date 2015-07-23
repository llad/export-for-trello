var bTrelloExporterLoaded=false;

function TrelloExportLoader()
{
      if(bTrelloExporterLoaded===true) return;
      setTimeout(function(){addExportLink();}, 500);
}

setInterval(function() {

  if($('.js-export-json').is(':visible') && bTrelloExporterLoaded===false) {
    setTimeout(function(){addExportLink();}, 500);
  }
  else {
    bTrelloExporterLoaded=false;
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
}
