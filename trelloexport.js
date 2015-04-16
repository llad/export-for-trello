/*!
 * TrelloExport
 * https://github.com/llad/export-for-trello
 *
 * Credit:
 * Started from: https://github.com/Q42/TrelloScrum
 */

/*jslint browser: true, devel: false*/

// Globals
var $,
    byteString,
    xlsx,
    ArrayBuffer,
    Uint8Array,
    Blob,
    saveAs;


// Variables
var $excel_btn,
    addInterval,
    columnHeadings = ['List', 'Title', 'Description', 'Points', 'Due', 'Members', 'Labels', 'Card #', 'Card URL'];

window.URL = window.webkitURL || window.URL;

function createExcelExport() {
    "use strict";
    // RegEx to find the points for users of TrelloScrum
    var pointReg = /[\(](\x3f|\d*\.?\d+)([\)])\s?/m;

	var boardExportURL = $('a.js-export-json').attr('href');
	//RegEx to extract Board ID
	var parts = /\/b\/(\w{8})\.json/.exec(boardExportURL);

	if(!parts) {
		alert("Board menu not open.");
		return;
	}

	var idBoard = parts[1];
	var apiURL = "https://trello.com/1/boards/" + idBoard + "?lists=open&cards=visible&card_attachments=cover&card_stickers=true&card_fields=badges%2Cclosed%2CdateLastActivity%2Cdesc%2CdescData%2Cdue%2CidAttachmentCover%2CidList%2CidBoard%2CidMembers%2CidShort%2Clabels%2CidLabels%2Cname%2Cpos%2CshortUrl%2CshortLink%2Csubscribed%2Curl&card_checklists=none&members=all&member_fields=fullName%2Cinitials%2CmemberType%2Cusername%2CavatarHash%2Cbio%2CbioData%2Cconfirmed%2Cproducts%2Curl%2Cstatus&membersInvited=all&membersInvited_fields=fullName%2Cinitials%2CmemberType%2Cusername%2CavatarHash%2Cbio%2CbioData%2Cconfirmed%2Cproducts%2Curl&checklists=none&organization=true&organization_fields=name%2CdisplayName%2Cdesc%2CdescData%2Curl%2Cwebsite%2Cprefs%2Cmemberships%2ClogoHash%2Cproducts&myPrefs=true&fields=name%2Cclosed%2CdateLastActivity%2CdateLastView%2CidOrganization%2Cprefs%2CshortLink%2CshortUrl%2Curl%2Cdesc%2CdescData%2Cinvitations%2Cinvited%2ClabelNames%2Cmemberships%2Cpinned%2CpowerUps%2Csubscribed";


    $.getJSON(apiURL, function (data) {

        var file = {
            worksheets: [[], []], // worksheets has one empty worksheet (array)
            creator: 'TrelloExport',
            created: new Date(),
            lastModifiedBy: 'TrelloExport',
            modified: new Date(),
            activeWorksheet: 0
        },

            // Setup the active list and cart worksheet
            w = file.worksheets[0],
            wArchived = file.worksheets[1],
            buffer,
            i,
            ia,
            blob,
            board_title;

        w.name = data.name.substring(0, 22);  // Over 22 chars causes Excel error, don't know why
        w.data = [];
        w.data.push([]);
        w.data[0] = columnHeadings;


        // Setup the archive list and cart worksheet
        wArchived.name = 'Archived ' + data.name.substring(0, 22);
        wArchived.data = [];
        wArchived.data.push([]);
        wArchived.data[0] = columnHeadings;

        // This iterates through each list and builds the dataset
        $.each(data.lists, function (key, list) {
            var list_id = list.id,
                listName = list.name;

            // tag archived lists
            if (list.closed) {
                listName = '[archived] ' + listName;
            }

            // Iterate through each card and transform data as needed
            $.each(data.cards, function (i, card) {
                if (card.idList === list_id) {
                    var title = card.name,
                        parsed = title.match(pointReg),
                        points = parsed ? parsed[1] : '',
                        due = card.due || '',
                        memberIDs,
                        memberInitials = [],
                        labels = [],
                        d = new Date(due),
                        rowData = [],
                        rArch,
                        r;

                    title = title.replace(pointReg, '');

                    // tag archived cards
                    if (card.closed) {
                        title = '[archived] ' + title;
                    }

                    memberIDs = card.idMembers;
                    $.each(memberIDs, function (i, memberID) {
                        $.each(data.members, function (key, member) {
                            if (member.id === memberID) {
                                memberInitials.push(member.initials);
                            }
                        });
                    });

                    //Get all labels
                    $.each(card.labels, function (i, label) {
                        if (label.name) {
                            labels.push(label.name);
                        } else {
                            labels.push(label.color);
                        }

                    });

                    // Need to set dates to the Date type so xlsx.js sets the right datatype
                    if (due !== '') {
                        due = d;
                    }

                    rowData = [
                        listName,
                        title,
                        card.desc,
                        points,
                        due,
                        memberInitials.toString(),
                        labels.toString(),
                        card.idShort,
						card.shortUrl
                    ];

                    // Writes all closed items to the Archived tab
                    // Note: Trello allows open cards on closed lists
                    if (list.closed || card.closed) {
                        rArch = wArchived.data.push([]) - 1;
                        wArchived.data[rArch] = rowData;

                    } else {
                        r = w.data.push([]) - 1;
                        w.data[r] = rowData;
                    }
                }
            });
        });

        // We want just the base64 part of the output of xlsx.js
        // since we are not leveraging they standard transfer process.
        byteString = window.atob(xlsx(file).base64);
        buffer = new ArrayBuffer(byteString.length);
        ia = new Uint8Array(buffer);

        // write the bytes of the string to an ArrayBuffer
        for (i = 0; i < byteString.length; i += 1) {
            ia[i] = byteString.charCodeAt(i);
        }

        // create blob and save it using FileSaver.js
        blob = new Blob([ia], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        board_title = data.name;
        saveAs(blob, board_title + '.xlsx');
        $("a.pop-over-header-close-btn")[0].click();


    });

}


// Add a Export Excel button to the DOM and trigger export if clicked
function addExportLink() {
    "use strict";
    //alert('add');

    var $js_btn = $('a.js-export-json'); // Export JSON link

    // See if our Export Excel is already there
    if ($('.pop-over-list').find('.js-export-excel').length) {
        clearInterval(addInterval);
        return;
    }

    // The new link/button
    if ($js_btn.length) {
        $excel_btn = $('<a>')
            .attr({
                'class': 'js-export-excel',
                'href': '#',
                'target': '_blank',
                'title': 'Open downloaded file with Excel'
            })
            .text('Export Excel')
            .click(createExcelExport)
            .insertAfter($js_btn.parent())
            .wrap(document.createElement("li"));

    }
}


// on DOM load
$(function () {
    "use strict";
    // Look for clicks on the .js-share class, which is
    // the "Share, Print, Export..." link on the board header option list
    $(document).on('mouseup', ".js-share", function () {
        addInterval = setInterval(addExportLink, 500);
    });
});
