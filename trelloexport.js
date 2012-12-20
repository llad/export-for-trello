//what to do when DOM loads
$(function () {
    $('.js-share').live('mouseup', function () {
        setTimeout(checkExport);
    });
});

//for export
var $excel_btn, $excel_dl;
window.URL = window.webkitURL || window.URL;

function checkExport() {
    if ($('form').find('.js-export-excel').length) return;
    var $js_btn = $('a.js-export-json');


    if ($js_btn.length) $excel_btn = $('<a>')
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
            w.data[0] = ['List', 'Title', 'Description', 'Points', 'Due'];
            $.each(data.lists, function (key, list) {
                var list_id = list.id;
                
                $.each(data.cards, function (i, card) {
                if (card.idList == list_id) {
                    var title = card.name;
                    var parsed = title.match(pointReg);
                    var points = parsed ? parsed[1] : '';
                    title = title.replace(pointReg, '');
                    var due = card.due || '';
                    
                    /* Member Listing Under Construction
                    var memberIDs = card.idMembers;
                    var memberInitials = [];
                    $.each(data.members, function (key, member) {
                        if (member.id == memberIDs[i]) {
                            console.log(memberIDs[i]);
                            memberInitials.push(member.initials);
                        }
                    });
                    */
                        var r = w.data.push([]) - 1;
                        w.data[r] = [list.name,
                                    title,
                                    card.desc,
                                    points,
                                    due,
                                    //memberInitials.toString()
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