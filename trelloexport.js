/*!
 * TrelloExport
 * https://github.com/llad/trelloExport
 *
 * Credit:
 * Started from: https://github.com/Q42/TrelloScrum
 */

// Variables
var $excel_btn;
window.URL = window.webkitURL || window.URL;


// on DOM load
$(function () {
    
    // Look for clicks on the .js-share class, which is
    // the "Share, Print, Export..." link on the board header option list
    $('.js-share').live('mouseup', function () {
        setTimeout(addExportLink);
    });
});


// Add a Export Excel button to the DOM and trigger export if clicked
function addExportLink() { 
   
    var $js_btn = $('a.js-export-json'); // Export JSON link
    
    // See if our Export Excel is already there
    if ($('form').find('.js-export-excel').length) return;
    
    // The new link/button
    if ($js_btn.length) $excel_btn = $('<a>')
        .attr({
        class: 'js-export-excel',
        href: '#',
        target: '_blank',
        title: 'Open downloaded file with Excel'
    })
        .text('Export Excel')
        .click(createExcelExport)
        .insertAfter($js_btn.parent())
        .wrap(document.createElement("li"));
}

function createExcelExport() {

    // RegEx to find the points for users of TrelloScrum
    var pointReg = /[\(](\x3f|\d*\.?\d+)([\)])\s?/m;

    $.getJSON($('a.js-export-json').attr('href'), function (data) {
        var file = {
            worksheets: [[]], // worksheets has one empty worksheet (array)
            creator: 'TrelloExport',
            created: new Date(),
            lastModifiedBy: 'TrelloExport',
            modified: new Date(),
            activeWorksheet: 0
            },
            w = file.worksheets[0]; // cache current worksheet
            w.name = data.name;
            w.data = [];
            w.data.push([]);
            w.data[0] = ['List', 'Title', 'Description', 'Points', 'Due', 'Members'];
            $.each(data.lists, function (key, list) {
                var list_id = list.id;
                
                $.each(data.cards, function (i, card) {
                if (card.idList == list_id) {
                    var title = card.name;
                    var parsed = title.match(pointReg);
                    var points = parsed ? parsed[1] : '';
                    title = title.replace(pointReg, '');
                    var due = card.due || '';
                    
                    //Get all the Member IDs
                    var memberIDs = card.idMembers;
                    var memberInitials = [];
                    $.each(memberIDs, function (i, memberID){
                        $.each(data.members, function (key, member) {
                            if (member.id == memberID) {
                                memberInitials.push(member.initials);
                            }
                        });
                    });
                    
                    //Format Due date field
                    if (due !== '' ){
                        var d = new Date(due);
                        due = d;
                        //due = d.getMonth() + '/' + d.getDate() + '/' + d.getFullYear().toString().substring(2, 4);
                    }
                                        
                    var r = w.data.push([]) - 1;
                    w.data[r] = [list.name,
                                title,
                                card.desc,
                                points,
                                due,
                                memberInitials.toString()
                                ];
                    }
            });
        });

        byteString = window.atob(xlsx(file).base64);
        var buffer = new ArrayBuffer(byteString.length);
        var ia = new Uint8Array(buffer);
        
        // write the bytes of the string to an ArrayBuffer
        for (var i = 0; i < byteString.length; i++) {
            ia[i] = byteString.charCodeAt(i);
        }
        
        var blob = new Blob([ia], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        var board_title = data.name;
        saveAs(blob, board_title + '.xlsx');
        $("a.close-btn")[0].click();


    });

    return false;

}