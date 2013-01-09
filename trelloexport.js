/*!
 * TrelloExport
 * https://github.com/llad/trelloExport
 *
 * Credit:
 * Started from: https://github.com/Q42/TrelloScrum
 */

// Variables
var $excel_btn,
    columnHeadings = ['List', 'Title', 'Description', 'Points', 'Due', 'Members', 'Labels'];

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
            worksheets: [[],[]], // worksheets has one empty worksheet (array)
            creator: 'TrelloExport',
            created: new Date(),
            lastModifiedBy: 'TrelloExport',
            modified: new Date(),
            activeWorksheet: 0
            },
            
            // Setup the active list and cart worksheet
            w = file.worksheets[0]; 
            w.name = data.name;
            w.data = [];
            w.data.push([]);
            w.data[0] = columnHeadings;
            
            
            // Setup the archive list and cart worksheet            
            wArchived = file.worksheets[1]; 
            wArchived.name = 'Archived ' + data.name;
            wArchived.data = [];
            wArchived.data.push([]);
            wArchived.data[0] = columnHeadings;
            
            // This iterates through each list and builds the dataset                     
            $.each(data.lists, function (key, list) {
                var list_id = list.id;
                var listName = list.name;                                            
                
                // tag archived lists
                if (list.closed) {
                    listName = '[archived] ' + listName;
                }
                
                // Iterate through each card and transform data as needed
                $.each(data.cards, function (i, card) {
                if (card.idList == list_id) {
                    var title = card.name;
                    var parsed = title.match(pointReg);
                    var points = parsed ? parsed[1] : '';
                    title = title.replace(pointReg, '');
                    
                    // tag archived cards
                    if (card.closed) {
                        title = '[archived] ' + title;
                    }
                    
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
                    
                    //Get all labels
                    var labels = [];
                    $.each(card.labels, function (i, label){
                        if (label.name) {
                            labels.push(label.name);
                        }
                        else {
                            labels.push(label.color);
                        }

                    });
                    
                    // Need to set dates to the Date type so xlsx.js sets the right datatype
                    if (due !== '' ){
                        var d = new Date(due);
                        due = d;
                    }
                    
                    var rowData = [
                            listName,
                            title,
                            card.desc,
                            points,
                            due,
                            memberInitials.toString(),
                            labels.toString()
                            ];
                    
                    // Writes all closed items to the Archived tab
                    // Note: Trello allows open cards on closed lists
                    if (list.closed || card.closed) {
                        var rArch = wArchived.data.push([]) - 1;
                        wArchived.data[rArch] = rowData;
                                                                                
                    }
                    else {                                         
                        var r = w.data.push([]) - 1;
                        w.data[r] = rowData;
                    }
                }
            });
        });
        
        // We want just the base64 part of the output of xlsx.js
        // since we are not leveraging they standard transfer process.
        byteString = window.atob(xlsx(file).base64);
        var buffer = new ArrayBuffer(byteString.length);
        var ia = new Uint8Array(buffer);
        
        // write the bytes of the string to an ArrayBuffer
        for (var i = 0; i < byteString.length; i++) {
            ia[i] = byteString.charCodeAt(i);
        }
        
        // create blob and save it using FileSaver.js
        var blob = new Blob([ia], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        var board_title = data.name;
        saveAs(blob, board_title + '.xlsx');
        $("a.close-btn")[0].click();


    });

}