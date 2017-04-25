/*
 * TrelloExport
 *
 * A Chrome extension for Trello, that allows to export boards to Excel spreadsheets. And more to come.
 *
 * Forked by @trapias (Alberto Velo)
 *  https://github.com/trapias/trelloExport
 * From:
 *  https://github.com/llad/trelloExport
 * Started from:
 *  https://github.com/Q42/TrelloScrum
 *
 * = = = VERSION HISTORY = = =
 * Whatsnew for version 1.8.8:
    - export Trello Plus Spent and Estimate data
    - export checklists
    - export comments (see commentLimit, default 100)
    - export attachments
    - export votes
    - use updated version of xlsx.js, modified by me (escapeXML)
    - use updated version of jquery, 2.1.0
* Whatsnew for version 1.8.9:
    - added column Card #
    - added columns memberCreator, datetimeCreated, datetimeDone and memberDone pulling modifications from https://github.com/bmccormack/export-for-trello/blob/5b2b8b102b98ed2c49241105cb9e00e44d4e1e86/trelloexport.js
    - added linq.min.js library to support linq queries for the above modifications
* Whatsnew for version 1.9.0:
    - switched to SheetJS library to export to excel, cfr https://github.com/SheetJS/js-xlsx
    - unicode characters are now correctly exported
* Whatsnew for version 1.9.1:
    - fixed button loading
    - some code cleaning
* Whatsnew for version 1.9.2:
    - fixed blocking error when duedate specified
    - new button loading function
* Whatsnew for version 1.9.3:
    - restored archived cards sheet
* Whatsnew for version 1.9.4:
    - fix exporting when there are no archived cards
* Whatsnew for version 1.9.5:
    - code lint
    - ignore case in finding 'Done' lists (thanks https://disqus.com/by/AlvonsiusAlbertNainupu/)
* Whatsnew for version 1.9.6:
    - order checklist items by position (issue #4)
    - minor code changes
* Whatsnew for version 1.9.7:
    - fix issue #3 (copied comments missing in export)
* Whatsnew for version 1.9.8:
    - use trello api to get data
* Whatsnew for version 1.9.9:
    - MAXCHARSPERCELL limit to avoid import errors in Excel (see https://support.office.com/en-nz/article/Excel-specifications-and-limits-16c69c74-3d6a-4aaf-ba35-e6eb276e8eaa)
    - removed commentLimit, all comments are loaded (but attention to MAXCHARSPERCELL limit above, since comments go to a single cell)
    - growl notifications with jquery-growl http://ksylvest.github.io/jquery-growl/
* Whatsnew for version 1.9.10:
    - adapt inject script to modified Trello layout
* Whatsnew for version 1.9.11:
    - options dialog
    - export full board or choosen list(s) only
    - add who and when item was completed to checklist items (issue #5 at GitHub)
* Whatsnew for version 1.9.12:
    - bugfix: the previously used BoardID was kept when navigating to another board
* Whatsnew for version 1.9.13:
    - new 'DoneTime' column holding card completion time in days, hours, minutes and seconds as per ISO8601 (cfr https://en.wikipedia.org/wiki/ISO_8601)
    - name (prefix) of 'Done' lists is now configurable, default "Done"
    - larger options dialog
    - export multiple (selected) boards
    - export multiple (selected) cards in a list
* Whatsnew for v. 1.9.14:
     - handle multiple nameListDone
     - fix getMoveCardAction and getCreateCardAction to properly handle actions when exporting from multiple boards
     - added a new function (getCommentCardActions) to load comments with a dedicated query
     - format dates according to locale
     - new capability to filter exported lists by name when exporting multiple boards
* Whatsnew for v. 1.9.15:
    - finally fix comments and done calculation exporting: thanks @fepsch
* Whatsnew for v. 1.9.16:
    - new icon
    - investigating issues reported by @fepsch
* Whatsnew for v. 1.9.17:
    - finally fix exporting ALL comments per card: now loading comments per single card, way slower but assures all comments are exported
* Whatsnew for v. 1.9.18:
    - improve UI: better feedback message timing, yet still blocking UI during export due to sync ajax requests
    - removed data limit setting from options dialog - just use 1000, maximum allowed by Trello APIs
    - fix filename (YYYYMMDDhhmmss)
    - fix some UI issues
* Whatsnew for v. 1.9.19:
    - refactoring export flow
    - updated jQuery Growl to version 1.3.1
    - new markdown export mode
* Whatsnew for v. 1.9.20:
    - fixes due to Trello UI changes
* Whatsnew for v. 1.9.21:
    - some UI (CSS) improvements
    - improved options dialog, resetting options when switching export type
    - new columns for Excel export: 'Total Checklist items' and 'Completed Checklist items'
    - better checklists formatting for Excel export
    - export to HTML
* Whatsnew for v. 1.9.22:
    - fix improper .md encoding as per issue #8 https://github.com/trapias/TrelloExport/issues/8
    - new option to decide whether to export archived items
* Whatsnew for v. 1.9.23:
    - OPML export
    - add memberDone to markdown and HTML exports
* Whatsnew for v. 1.9.24:
    - checkboxes to enable/disable exporting of comments, checklist items and attachments
    - new option to export checklist items to rows, for Excel only
* Whatsnew for v. 1.9.25:
    - paginate cards loading so to be able to load all cards, even if exceeding the API limit of 1000 records per call
* Whatsnew for v. 1.9.26 (thanks chris https://github.com/collisdigital):
    - export points estimate and consumed from Card titles based on Scrum for Trello
    - improved regex for Trello Plus estimate/spent in card titles
* Whatsnew for v. 1.9.27:
    - fix ajax.fail functions
    - fix loading boards when current board does not belong to any organization
* Whatsnew for v. 1.9.28:
    - fix cards loading: something is broken with the paginated loading introduced with 1.9.25; to be further investigated
* Whatsnew for v. 1.9.29:
    - fix bug using user fullName, might not be available (thanks Natalia L.)
    - new css to format HTML exported files
* Whatsnew for v. 1.9.30:
    - fix 1.9.29 beta
    - finalize new css for HTML exports
* Whatsnew for v. 1.9.31:
    - fix due date format in Excel export (issue #12)
    - fix missing export of archived cards (issue #13)
* Whatsnew for v. 1.9.32:
    - hopefully fixed bug with member fullName reading
    - new option to export labels and members to Excel rows, like already available for checklist items (issue #15 https://github.com/trapias/TrelloExport/issues/15)
    - new option to show attached images inline for Markdown and HTML exports (issue #16 https://github.com/trapias/TrelloExport/issues/16)
* Whatsnew for v. 1.9.33:
    - new data field dateLastActivity exported (issue #18)
    - new data field numberOfComments exported (issue #19)
    - new option to choose which columns to export to Excel (issue #17)
* Whatsnew for v. 1.9.34:
    - only show columns chooser for Excel exports
    - can now set a custom css for HTML export
    - can now check/uncheck all columns to export
* Whatsnew for v. 1.9.35:
    - fix Trello header css height
* Whatsnew for v. 1.9.36:
    - filter by list name, card name or label name
    - help tooltips
*/
var VERSION = '1.9.36';

/**
 * http://stackoverflow.com/questions/784586/convert-special-characters-to-html-in-javascript
 * (c) 2012 Steven Levithan <http://slevithan.com/>
 * MIT license
 */
if (!String.prototype.codePointAt) {
    String.prototype.codePointAt = function(pos) {
        pos = isNaN(pos) ? 0 : pos;
        var str = String(this),
            code = str.charCodeAt(pos),
            next = str.charCodeAt(pos + 1);
        // If a surrogate pair
        if (0xD800 <= code && code <= 0xDBFF && 0xDC00 <= next && next <= 0xDFFF) {
            return ((code - 0xD800) * 0x400) + (next - 0xDC00) + 0x10000;
        }
        return code;
    };
}

/**
 * Encodes special html characters
 * @param string
 * @return {*}
 */
function html_encode(string) {
    var ret_val = '';
    for (var i = 0; i < string.length; i++) {
        if (string.codePointAt(i) > 127) {
            ret_val += '&#' + string.codePointAt(i) + ';';
        } else {
            ret_val += string.charAt(i);
        }
    }
    return ret_val;
}

String.prototype.replaceAll = function(search, replacement) {
    var target = this;
    return target.replace(new RegExp(search, 'g'), replacement);
};

function escape4XML(s) {
    s = s.replaceAll('"', '&quot;');
    s = s.replaceAll("'", '&apos;');
    s = s.replaceAll('<', '&lt;');
    s = s.replaceAll('>', '&gt;');
    s = s.replaceAll('&', '&amp;');
    s = html_encode(s);
    return s;
}

var $,
    byteString,
    xlsx,
    ArrayBuffer,
    Uint8Array,
    // Blob,
    // saveAs,
    actionsCreateCard = [],
    actionsMoveCard = [],
    actionsCommentCard = [],
    idBoard,
    nProcessedBoards = 0,
    nProcessedLists = 0,
    nProcessedCards = 0,
    $excel_btn,
    dataLimit = 1000, // limit the number of items retrieved from Trello (1000 is max allowed by Trello API server)
    MAXCHARSPERCELL = 32767,
    exportlists = [],
    exportboards = [],
    exportcards = [],
    nameListDone = "Done",
    filterListsNames = [],
    pageSize = 1000;

function sheet_from_array_of_arrays(data, opts) {
    // console.log('sheet_from_array_of_arrays ' + data);
    var ws = {};
    var range = {
        s: {
            c: 10000000,
            r: 10000000
        },
        e: {
            c: 0,
            r: 0
        }
    };
    for (var R = 0; R != data.length; ++R) {
        for (var C = 0; C != data[R].length; ++C) {
            if (range.s.r > R) range.s.r = R;
            if (range.s.c > C) range.s.c = C;
            if (range.e.r < R) range.e.r = R;
            if (range.e.c < C) range.e.c = C;
            var cell = {
                v: data[R][C]
            };
            if (cell.v === null) continue;
            var cell_ref = XLSX.utils.encode_cell({
                c: C,
                r: R
            });

            if (typeof cell.v === 'number') cell.t = 'n';
            else if (typeof cell.v === 'boolean') cell.t = 'b';
            else if (cell.v instanceof Date) {
                cell.t = 'n';
                cell.z = XLSX.SSF._table[14];
                cell.v = datenum(cell.v);
            } else cell.t = 's';

            ws[cell_ref] = cell;
        }
    }
    if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
    return ws;
}

function Workbook() {
    if (!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}

function dd(s) {
    return ('0' + s).slice(-2);
}

function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}

if (typeof String.prototype.startsWith != 'function') {
    String.prototype.startsWith = function(str) {
        return this.indexOf(str) === 0;
    };
}

function getCommentCardActions(boardID, idCard) {
    for (var n = 0; n < actionsCommentCard.length; n++) {
        if (actionsCommentCard[n].card === idCard) {
            var query = Enumerable.From(actionsCommentCard[n].data)
                .Where(function(x) {
                    if (x.data.card) {
                        return x.data.card.id == idCard;
                    }
                })
                .OrderByDescending(function(x) {
                    return x.date;
                })
                .ToArray();
            return query.length > 0 ? query : false;
        }
    }
    $.ajax({
        url: 'https://trello.com/1/card/' + idCard + '/actions?filter=commentCard,copyCommentCard&limit=' + dataLimit,
        // url:'https://trello.com/1/boards/' + boardID + '/actions?filter=commentCard,copyCommentCard&limit=' + dataLimit,
        dataType: 'json',
        async: false,
        success: function(actionsData) {
            var a = {
                card: idCard,
                data: actionsData,
            };
            actionsCommentCard.push(a);
        }
    });
    var selectedActions = null;
    // get the right actions for board
    for (var i = 0; i < actionsCommentCard.length; i++) {
        if (actionsCommentCard[i].card === idCard) {
            selectedActions = actionsCommentCard[i].data;
            break;
        }
    }
    var query2 = Enumerable.From(selectedActions)
        .Where(function(x) {
            if (x.data.card) {
                return x.data.card.id == idCard;
            }
        })
        .OrderByDescending(function(x) {
            return x.date;
        })
        .ToArray();
    return query2.length > 0 ? query2 : false;
}

function getCreateCardAction(boardID, idCard) {
    for (var n = 0; n < actionsCreateCard.length; n++) {
        if (actionsCreateCard[n].board === boardID) {
            var query = Enumerable.From(actionsCreateCard[n].data)
                .Where(function(x) {
                    if (x.data.card) {
                        return x.data.card.id == idCard;
                    }
                })
                .ToArray();
            return query.length > 0 ? query[0] : false;
        }
    }
    $.ajax({
        url: 'https://trello.com/1/boards/' + boardID + '/actions?filter=createCard&limit=' + dataLimit,
        dataType: 'json',
        async: false,
        success: function(actionsData) {
            var a = {
                board: boardID,
                data: actionsData,
            };
            actionsCreateCard.push(a);
        }
    });
    var selectedActions = null;
    // get the right actions for board
    for (var i = 0; i < actionsCreateCard.length; i++) {
        if (actionsCreateCard[i].board === boardID) {
            selectedActions = actionsCreateCard[i].data;
            break;
        }
    }
    var query2 = Enumerable.From(selectedActions)
        .Where(function(x) {
            if (x.data.card) {
                return x.data.card.id == idCard;
            }
        })
        .ToArray();
    return query2.length > 0 ? query2[0] : false;
}

function getMoveCardAction(boardID, idCard, nameList) {
    for (var n = 0; n < actionsMoveCard.length; n++) {
        if (actionsMoveCard[n].board === boardID) {
            var query = Enumerable.From(actionsMoveCard[n].data)
                .Where(function(x) {
                    if (x.data.card && x.data.listAfter) {
                        return x.data.card.id == idCard && x.data.listAfter.name == nameList;
                    }
                })
                .OrderByDescending(function(x) {
                    return x.date;
                })
                .ToArray();
            return query.length > 0 ? query[0] : false;
        }
    }
    $.ajax({
        url: 'https://trello.com/1/boards/' + boardID + '/actions?filter=updateCard&limit=' + dataLimit,
        dataType: 'json',
        async: false,
        success: function(actionsData) {
            var a = {
                board: boardID,
                data: actionsData,
            };
            actionsMoveCard.push(a);
        }
    });
    var selectedActions = null;
    // get the right actions for board
    for (var i = 0; i < actionsMoveCard.length; i++) {
        if (actionsMoveCard[i].board === boardID) {
            selectedActions = actionsMoveCard[i].data;
            break;
        }
    }
    var query2 = Enumerable.From(selectedActions)
        .Where(function(x) {
            if (x.data.card && x.data.listAfter) {
                return x.data.card.id == idCard && x.data.listAfter.name == nameList;
            }
        })
        .OrderByDescending(function(x) {
            return x.date;
        })
        .ToArray();
    return query2.length > 0 ? query2[0] : false;
}

function searchupdateCheckItemStateOnCardAction(checkitemid, actions) {
    var completedObject = {};
    $.each(actions, function(j, action) {
        if (action.type == 'updateCheckItemStateOnCard') {
            if (action.data.checkItem.id == checkitemid) {
                var d = new Date(action.date);
                var sActionDate = d.toLocaleDateString() + ' ' + d.toLocaleTimeString();
                completedObject.date = sActionDate;
                if (action.memberCreator !== undefined) {
                    completedObject.by = (action.memberCreator.fullName !== undefined ? action.memberCreator.fullName : action.memberCreator.username);
                } else {
                    completedObject.by = '';
                }
                return completedObject;
            }
        }
    });
    return completedObject;
}

function TrelloExportOptions() {

    exportboards = []; // reset
    exportlists = [];
    exportcards = [];
    filterListsNames = [];
    nProcessedBoards = 0;
    nProcessedLists = 0;
    nProcessedCards = 0;
    var $js_btn = $('a.js-export-json'); // Export JSON link
    var boardExportURL = $js_btn.attr('href');
    var parts = /\/b\/(\w{8})\.json/.exec(boardExportURL); // extract board id
    if (!parts) {
        $.growl.error({
            title: "TrelloExport",
            message: "Board menu not open?",
            fixed: true
        });
        return;
    }
    idBoard = parts[1];

    var columnHeadings = [
        'Board', 'List', 'Card #', 'Title', 'Link', 'Description',
        'Total Checklist items', 'Completed Checklist items', 'Checklists',
        'NumberOfComments', 'Comments', 'Attachments', 'Votes', 'Spent', 'Estimate',
        'Points Estimate', 'Points Consumed', 'Created', 'CreatedBy', 'LastActivity', 'Due',
        'Done', 'DoneBy', 'DoneTime', 'Members', 'Labels'
    ];

    // https://github.com/davidstutz/bootstrap-multiselect
    var options = [];
    for (var x = 0; x < columnHeadings.length; x++) {
        var o = '<option value="' + columnHeadings[x] + '" selected="true">' + columnHeadings[x] + '</option>';
        options.push(o);
    }

    var sDialog = '<table id="optionslist">' +
        '<tr><td><span data-toggle="tooltip" data-placement="right" data-container="body" title="Choose the type of file you want to export">Export to</span></td><td><select id="exportmode"><option value="XLSX">Excel</option><option value="MD">Markdown</option><option value="HTML">HTML</option><option value="OPML">OPML</option></select></td></tr>' +
        '<tr><td><span data-toggle="tooltip" data-placement="right" data-container="body" title="Check all the kinds of items you want to export">Export:</span></td><td><input type="checkbox" id="exportArchived" title="Export archived cards">Archived cards ' +
        '<input type="checkbox" id="comments" title="Export comments">Comments<br/><input type="checkbox" id="checklists" title="Export checklists">Checklists <input type="checkbox" id="attachments" title="Export attachments">Attachments</td></tr>' +
        '<tr id="cklAsRowsRow"><td><span data-toggle="tooltip" data-placement="right" data-container="body" title="Create one Excel row per each card, checklist item, label or card member">One row per each:</span></td><td><input type="radio" id="cardsAsRows" checked name="asrows" value="0"> <label for="cardsAsRows" >Card</label>  <input type="radio" id="cklAsRows" name="asrows" value="1"> <label for="cklAsRows">Checklist item</label>  <input type="radio" id="lblAsRows" name="asrows" value="2"> <label for="lblAsRows">Label</label>  <input type="radio" id="membersAsRows" name="asrows" value="3"> <label for="membersAsRows">Member</label>  </td></tr>' +
        '<tr id="xlsColumns">' +
        '<td><span data-toggle="tooltip" data-placement="right" data-container="body" title="Choose columns to be exported to Excel">Export columns</span></td>' +
        '<td><select multiple="multiple" id="selectedColumns">' + options.join('') + '</select></td>' +
        '</tr>' +
        '<tr id="ckHTMLCardInfoRow" style="display:none"><td><span data-toggle="tooltip" data-placement="right" data-container="body" title="Set options for the target HTML">Options:</span></td><td><input type="checkbox" checked id="ckHTMLCardInfo" title="Export card info"> Export card info (created, createdby) <br/><input type="checkbox" checked id="chkHTMLInlineImages" title="Show attachment images"> Show attachment images' +
        '<br/>Stylesheet: <input id="trelloExportCss" type="text" name="css" value="http://trapias.github.io/assets/TrelloExport/default.css"> ' + '</td></tr>' +
        '<tr><td><span data-toggle="tooltip" data-placement="right" data-container="body" title="Set the List name prefix used to recognize your completed lists. See http://trapias.github.io/blog/trelloexport-1-9-13">Done lists name:</span></td><td><input type="text" size="4" name="setnameListDone" id="setnameListDone" value="' + nameListDone + '"  placeholder="Set prefix or leave empty" /></td></tr>' +
        '<tr><td><span data-toggle="tooltip" data-placement="right" data-container="body" title="Choose what data to export">Type of export:</span></td><td><select id="exporttype"><option value="board">Current Board</option><option value="list">Select Lists in current Board</option><option value="boards">Multiple Boards</option><option value="cards">Select cards in a list</option></select></td></tr>' +
        '<tr><td><span data-toggle="tooltip" data-placement="right" data-container="body" title="Only include items whose name starts with the specified prefix">Filter:</span></td><td>' +
        '<select id="filterMode"><option value="List">On List name</option><option value="Label">On Label name</option><option value="Card">On card name</option></select>' +
        '<input type="text" size="4" name="filterListsNames" id="filterListsNames" value="" placeholder="Set prefix or leave empty" /></td></tr></table>';

    var dlgReady = new Promise(
        function(resolve, reject) {

            // open options dialog to configure & launch export
            $.Zebra_Dialog(sDialog, {
                title: 'TrelloExport ' + VERSION,
                type: false,
                'buttons': [{
                    caption: 'Export',
                    callback: function() {

                        nameListDone = $('#setnameListDone').val();
                        var mode = $('#exportmode').val();
                        var sfilterListsNames, filters, bexportArchived, bExportComments, bExportChecklists, bExportAttachments, iExcelItemsAsRows, bckHTMLCardInfo, bchkHTMLInlineImages;
                        bexportArchived = $('#exportArchived').is(':checked');
                        bExportComments = $('#comments').is(':checked');
                        bExportChecklists = $('#checklists').is(':checked');
                        bExportAttachments = $('#attachments').is(':checked');
                        iExcelItemsAsRows = 0;
                        iExcelItemsAsRows = $('input[name=asrows]:checked').val();
                        bckHTMLCardInfo = $('#ckHTMLCardInfo').is(':checked');
                        bchkHTMLInlineImages = $('#chkHTMLInlineImages').is(':checked');

                        if (!bExportChecklists && iExcelItemsAsRows.toString() === '1') {
                            // checklist items as rows only available if checklists are exported
                            iExcelItemsAsRows = 0;
                        }
                        // export type
                        var sexporttype = $('#exporttype').val();
                        switch (sexporttype) {
                            case 'list':
                                if ($('#choosenlist').length <= 0) {
                                    // console.log('wait for lists to load');
                                    return false;
                                } else {
                                    $('#choosenlist > option:selected').each(function() {
                                        exportlists.push($(this).val());
                                    });
                                }
                                break;

                            case 'boards':
                                if ($('#choosenboards').length <= 0) {
                                    // console.log('wait for lists to load');
                                    return false;
                                } else {
                                    sfilterListsNames = $('#filterListsNames').val();
                                    if (sfilterListsNames.trim() !== '') {
                                        // parse list name filters
                                        filters = sfilterListsNames.split(',');
                                        for (var nd = 0; nd < filters.length; nd++) {
                                            filterListsNames.push(filters[nd].toString().trim());
                                        }
                                    }
                                    $('#choosenboards > option:selected').each(function() {
                                        exportboards.push($(this).val());
                                    });
                                }
                                break;

                            case 'cards':
                                // choosenCards
                                if ($('#choosenCards').length <= 0) {
                                    // console.log('wait for cards to load');
                                    return false;
                                } else {
                                    $('#choosenCards > option:selected').each(function() {
                                        exportcards.push($(this).val());
                                    });
                                }
                                break;

                            case 'board':
                                sfilterListsNames = $('#filterListsNames').val();
                                if (sfilterListsNames.trim() !== '') {
                                    // parse list name filters
                                    filters = sfilterListsNames.split(',');
                                    for (var nd2 = 0; nd2 < filters.length; nd2++) {
                                        filterListsNames.push(filters[nd2].toString().trim());
                                    }
                                }
                                $('#choosenboards > option:selected').each(function() {
                                    exportboards.push($(this).val());
                                });
                                break;

                            default:
                                break;
                        }

                        var allColumns = $('#selectedColumns option');
                        var selectedColumns = [];
                        var selectedOptions = $('#selectedColumns option:selected');
                        selectedOptions.each(function() {
                            selectedColumns.push(this.value);
                        });

                        var css = $('#trelloExportCss').val();

                        // filterMode
                        var filterMode = $('#filterMode').val();

                        // launch export
                        setTimeout(function() {
                            loadData(mode, bexportArchived, bExportComments, bExportChecklists, bExportAttachments, iExcelItemsAsRows, bckHTMLCardInfo, bchkHTMLInlineImages, allColumns, selectedColumns, css, filterMode);
                        }, 500);
                        return true;
                    }
                }, {
                    caption: 'Cancel',
                    callback: function() {
                            return;
                        } // close dialog
                }]
            });

            resolve('dlgready');
        });

    dlgReady.then(function() {

        $('[data-toggle="tooltip"]').tooltip();

        $('#selectedColumns').multiselect({ includeSelectAllOption: true });

        $('#exporttype').on('change', function() {
            var sexporttype = $('#exporttype').val();
            var sSelect;
            resetOptions();

            switch (sexporttype) {
                case 'list':
                    // get a list of all lists in board and let user choose which to export
                    sSelect = getalllistsinboard();
                    $('#optionslist').append('<tr><td>Select one or more Lists</td><td><select multiple id="choosenlist">' + sSelect + '</select></td></tr>');
                    break;
                case 'board':
                    $('#optionslist').append('<tr><td>Filter lists by name:</td><td><input type="text" size="4" name="filterListsNames" id="filterListsNames" value="" placeholder="Set prefix or leave empty" /></td></tr>');
                    break;
                case 'boards':
                    // get a list of all boards
                    sSelect = getallboards();
                    $('#optionslist').append('<tr><td>Select one or more Boards</td><td><select multiple id="choosenboards">' + sSelect + '</select></td></tr>');
                    $('#optionslist').append('<tr><td>Filter lists by name:</td><td><input type="text" size="4" name="filterListsNames" id="filterListsNames" value="" placeholder="Set prefix or leave empty" /></td></tr>');
                    break;
                case 'cards':
                    // get a list of all lists in board and let user choose which to export
                    sSelect = getalllistsinboard();
                    $('#optionslist').append('<tr><td>Select one List</td><td><select id="choosenSinglelist"><option value="">Select a list</option>' + sSelect + '</select></td></tr>');

                    $('#choosenSinglelist').on('change', function() {
                        $('#choosenCards').parent().parent().remove();
                        var selectedList = $('#choosenSinglelist').val();
                        exportlists = [];
                        exportlists.push(selectedList);
                        // get card in list
                        sSelect = getallcardsinlist(selectedList);
                        $('#optionslist').append('<tr><td>Select one or more cards</td><td><select multiple id="choosenCards">' + sSelect + '</select></td></tr>');
                    });
                    break;
                default:
                    break;
            }

        });

        $('#exportmode').on('change', function() {
            var mode = $('#exportmode').val();
            $('#cklAsRowsRow').hide();
            $('#ckHTMLCardInfoRow').hide();
            $('#xlsColumns').hide();

            switch (mode) {
                case 'XLSX':
                    $('#cklAsRowsRow').show();
                    $('#xlsColumns').show();
                    break;
                case 'HTML':
                case 'MD':
                    $('#ckHTMLCardInfoRow').show();
                    break;
                default:

                    break;
            }

        });

        $('#cklAsRows').attr('disabled', true); //  default
        $('#checklists').on('change', function() {
            var ecl = $('#checklists').is(':checked');
            if (!ecl) {
                $('#cklAsRows').attr('disabled', true);

            } else {
                $('#cklAsRows').attr('disabled', false);
            }
        });

        $('input[name=asrows]').change(function() {
            setColumnHeadings(this.value);
        });

        // todo: show progress
        //

    });

    return; // close dialog
}

function setColumnHeadings(asrowsMode) {
    var columnHeadings = [];
    switch (Number(asrowsMode)) {
        case 1: // checklist item
            columnHeadings = [
                'Board', 'List', 'Card #', 'Title', 'Link', 'Description',
                'Total Checklist items', 'Completed Checklist items', 'Checklist',
                'Checklist item', 'Completed', 'DateCompleted', 'CompletedBy',
                'NumberOfComments', 'Comments', 'Attachments', 'Votes', 'Spent', 'Estimate',
                'Points Estimate', 'Points Consumed', 'Created', 'CreatedBy', 'LastActivity', 'Due',
                'Done', 'DoneBy', 'DoneTime', 'Members', 'Labels'
            ];
            break;
        case 2: // label
            columnHeadings = [
                'Board', 'List', 'Card #', 'Title', 'Link', 'Description',
                'Total Checklist items', 'Completed Checklist items', 'Checklists',
                'NumberOfComments', 'Comments', 'Attachments', 'Votes', 'Spent', 'Estimate',
                'Points Estimate', 'Points Consumed', 'Created', 'CreatedBy', 'LastActivity', 'Due',
                'Done', 'DoneBy', 'DoneTime', 'Members', 'Label'
            ];
            break;
        case 3: // member
            console.log('MEMBER');
            columnHeadings = [
                'Board', 'List', 'Card #', 'Title', 'Link', 'Description',
                'Total Checklist items', 'Completed Checklist items', 'Checklists',
                'NumberOfComments', 'Comments', 'Attachments', 'Votes', 'Spent', 'Estimate',
                'Points Estimate', 'Points Consumed', 'Created', 'CreatedBy', 'LastActivity', 'Due',
                'Done', 'DoneBy', 'DoneTime', 'Member', 'Labels'
            ];
            break;
        default:
            // card
            columnHeadings = [
                'Board', 'List', 'Card #', 'Title', 'Link', 'Description',
                'Total Checklist items', 'Completed Checklist items', 'Checklists',
                'NumberOfComments', 'Comments', 'Attachments', 'Votes', 'Spent', 'Estimate',
                'Points Estimate', 'Points Consumed', 'Created', 'CreatedBy', 'LastActivity', 'Due',
                'Done', 'DoneBy', 'DoneTime', 'Members', 'Labels'
            ];
            break;
    }

    var options = [];
    for (var x = 0; x < columnHeadings.length; x++) {
        var o = '<option value="' + columnHeadings[x] + '" selected="true">' + columnHeadings[x] + '</option>';
        options.push(o);
    }

    $('#selectedColumns').multiselect('destroy')
        .find('option')
        .remove()
        .end()
        .append(options.join(''))
        .multiselect({ includeSelectAllOption: true });
}

function resetOptions() {
    $('#choosenlist').parent().parent().remove();
    $('#choosenboards').parent().parent().remove();
    $('#choosenCards').parent().parent().remove();
    $('#choosenSinglelist').parent().parent().remove();
    $('#filterListsNames').parent().parent().remove();
}

function getalllistsinboard() {
    var apiURL = "https://trello.com/1/boards/" + idBoard + "?lists=all&cards=none";
    var sHtml = "";

    $.ajax({
            url: apiURL,
            async: false,
        })
        .done(function(data) {
            // console.log('DATA:' + JSON.stringify(data));
            $.each(data.lists, function(key, list) {
                var list_id = list.id;
                var listName = list.name;
                if (!list.closed) {
                    sHtml += '<option value="' + list_id + '">' + listName + '</option>';
                } else {
                    sHtml += '<option value="' + list_id + '">' + listName + ' [Archived]</option>';
                }
            });
        })
        .fail(function(jqXHR, textStatus, errorThrown) {
            console.log("getalllistsinboard error!!!");
            $.growl.error({
                title: "TrelloExport",
                message: jqXHR.statusText + ' ' + jqXHR.status + ': ' + jqXHR.responseText,
                fixed: true
            });
        })
        .always(function() {
            // console.log("complete");
        });

    return sHtml;
}

function getorganizationid() {
    var apiURL = "https://trello.com/1/boards/" + idBoard + '?lists=none';
    var orgID = "";

    $.ajax({
            url: apiURL,
            async: false,
        })
        .done(function(data) {
            //console.log('DATA:' + JSON.stringify(data));
            orgID = data.idOrganization;
        })
        .fail(function(jqXHR, textStatus, errorThrown) {
            console.log("getorganizationid error!!!");
            $.growl.error({
                title: "TrelloExport",
                message: jqXHR.statusText + ' ' + jqXHR.status + ': ' + jqXHR.responseText,
                fixed: true
            });
        })
        .always(function() {
            // console.log("complete");
        });

    return orgID;
}

function getallboards() {

    var orgID = getorganizationid();

    // GET /1/organizations/[idOrg or name]/boards
    var apiURL = "https://trello.com/1/organizations/" + orgID + "/boards?lists=none&fields=name,idOrganization,closed";
    var sHtml = "";

    if (orgID === null) {
        // current board outside any organization, get all boards
        apiURL = "https://trello.com/1/members/me/boards?lists=none&fields=name,idOrganization,closed";
    }

    $.ajax({
            url: apiURL,
            async: false,
        })
        .done(function(data) {
            for (var i = 0; i < data.length; i++) {
                var board_id = data[i].id;
                var boardName = data[i].name;
                if (!data[i].closed) {
                    sHtml += '<option value="' + board_id + '">' + boardName + '</option>';
                } else {
                    sHtml += '<option value="' + board_id + '">' + boardName + ' [Archived]</option>';
                }
            }
        })
        .fail(function(jqXHR, textStatus, errorThrown) {
            console.log("getallboards error: " + jqXHR.statusText + ' ' + jqXHR.status + ': ' + jqXHR.responseText);
            $.growl.error({
                title: "TrelloExport",
                message: jqXHR.statusText + ' ' + jqXHR.status + ': ' + jqXHR.responseText,
                fixed: true
            });
        })
        .always(function() {
            // console.log("complete");
        });

    return sHtml;
}

function getBoardName(id) {

    var apiURL = "https://trello.com/1/boards/" + id + "?fields=name";
    var sName = "";

    $.ajax({
            url: apiURL,
            async: false,
        })
        .done(function(data) {
            sName = data.name;
        })
        .fail(function(jqXHR, textStatus, errorThrown) {
            console.log("getBoardName error!!!");
            $.growl.error({
                title: "TrelloExport",
                message: jqXHR.statusText + ' ' + jqXHR.status + ': ' + jqXHR.responseText,
                fixed: true
            });
        })
        .always(function() {
            // console.log("complete");
        });

    return sName;
}

function getallcardsinlist(listid) {
    // GET /1/lists/[idList]/cards?fields=idShort
    var apiURL = "https://trello.com/1/lists/" + listid + "/cards?fields=idShort,name,closed";
    var sHtml = "";

    $.ajax({
            url: apiURL,
            async: false,
        })
        .done(function(data) {
            // console.log('Got cards: ' + JSON.stringify(data));
            for (var i = 0; i < data.length; i++) {
                var card_id = data[i].id;
                var cardName = data[i].name;
                if (!data[i].closed) {
                    sHtml += '<option value="' + card_id + '">' + cardName + '</option>';
                } else {
                    sHtml += '<option value="' + card_id + '">' + cardName + ' [Archived]</option>';
                }
            }
        })
        .fail(function(jqXHR, textStatus, errorThrown) {
            console.log("getallcardsinlist error!!!");
            $.growl.error({
                title: "TrelloExport",
                message: jqXHR.statusText + ' ' + jqXHR.status + ': ' + jqXHR.responseText,
                fixed: true
            });
        })
        .always(function() {
            // console.log("complete");
        });

    return sHtml;
}

function extractFloat(str, regex, groupIndex) {
    var value = ''
    var match = str.match(regex);
    if (match !== null) {
        value = parseFloat(match[groupIndex]);
        if (!!value === false) {
            value = '';
        }
    }
    return value;
}


function loadData(exportFormat, bexportArchived, bExportComments, bExportChecklists, bExportAttachments, iExcelItemsAsRows, bckHTMLCardInfo, bchkHTMLInlineImages, allColumns, selectedColumns, css, filterMode) {
    console.log('TrelloExport loading data for export format: ' + exportFormat + '...');

    var promLoadData = new Promise(
        function(resolve, reject) {
            $.growl({
                title: "TrelloExport",
                message: "Loading data, please wait..."
            });

            setTimeout(function() {
                resolve();
            }, 100);
        });

    promLoadData.then(function() {

        var jsonComputedCards = [];

        // RegEx to find the Trello Plus Spent and Estimate (spent/estimate) in card titles
        var SpentEstRegex = /(\(([0-9]+(\.[0-9]+)?)(\/?([0-9]+(\.[0-9]+)?))?\))/;
        var PointsEstRegex = /(\(([0-9]+(\.[0-9]+)?)\))/;
        var PointsConRegex = /(\[([0-9]+(\.[0-9]+)?)\])/;

        /*
            get data via Trello API instead of Board's json
        */

        if (exportboards.length === 0) {
            // export just current board
            exportboards.push(idBoard);
        }

        // loop boards
        for (var iBoard = 0; iBoard < exportboards.length; iBoard++) {
            console.log('Export board ' + exportboards[iBoard]);
            idBoard = exportboards[iBoard];
            var boardName = getBoardName(idBoard);

            var readCards = 0;

            var apiURL = "https://trello.com/1/boards/" + idBoard + "/lists?fields=name,closed"; //"?lists=all&cards=all&card_fields=all&card_checklists=all&members=all&member_fields=all&membersInvited=all&checklists=all&organization=true&organization_fields=all&fields=all&actions=commentCard%2CcopyCommentCard%2CupdateCheckItemStateOnCard&card_attachments=true";
            $.ajax({
                    url: apiURL,
                    async: false,
                })
                .done(function(data) {

                    console.log('DONE: got ' + data.length + ' lists');

                    // if (!bexportArchived) {
                    //     if (data.closed) {
                    //         console.log('Skip archived Board ' + data.id);
                    //         return true;
                    //     }
                    // }

                    // This iterates through each list and builds the dataset
                    $.each(data, function(key, list) {

                        var sBefore = ''; // reset for each list

                        var list_id = list.id;
                        var listName = list.name;

                        if (!bexportArchived) {
                            if (list.closed) {
                                console.log('skip archived list ' + listName);
                                return true;
                            }
                        }

                        if (exportlists.length > 0) {

                            if ($.inArray(list_id, exportlists) === -1) {
                                console.log('skip list ' + listName);
                                return true;
                            }
                        }

                        // 1.9.14: filter lists by name
                        var accept = true;
                        if (filterListsNames.length > 0 && filterMode === 'List') {
                            for (var y = 0; y < filterListsNames.length; y++) {
                                if (!listName.toLowerCase().startsWith(filterListsNames[y].trim().toLowerCase())) {
                                    accept = false;
                                } else {
                                    accept = true;
                                    break;
                                }
                            }
                        }
                        if (!accept) {
                            console.log('skipping list ' + listName);
                            return true;
                        }

                        console.log('processing list ' + listName);
                        nProcessedLists++;

                        // tag archived lists
                        if (list.closed) {
                            listName = '[archived] ' + listName;
                        }

                        var exportedCardsIDs = [];

                        do {
                            readCards = 0;
                            $.ajax({
                                    url: "https://trello.com/1/lists/" + list_id + "/cards?filter=all&fields=all&checklists=all&members=true&member_fields=all&membersInvited=all&organization=true&organization_fields=all&actions=commentCard%2CcopyCommentCard%2CupdateCheckItemStateOnCard&attachments=true&limit=" + pageSize + "&before=" + sBefore,
                                    async: false,
                                })
                                .done(function(datacards) {

                                    readCards = datacards.length;
                                    //console.log('evaluating ' + datacards.length + ' cards');

                                    // Iterate through each card and transform data as needed
                                    $.each(datacards, function(i, card) {

                                        // console.log('=-=-=-=-=-=-=-=>>>>> I = ' + i);

                                        if (sBefore === card.id) {
                                            // console.log('ALREADY USED BEFORE VALUE ' + sBefore);
                                            readCards = -1;
                                            return;
                                        }

                                        if (i === 0) {
                                            sBefore = card.id;
                                        }

                                        if (card.idList == list_id && $.inArray(card.id, exportedCardsIDs) === -1) {

                                            if (!bexportArchived) {
                                                if (card.closed) {
                                                    console.log('Skip archived card ' + card.name);
                                                    return true;
                                                }
                                            }

                                            //export selected cards only option
                                            if (exportcards.length > 0) {
                                                if ($.inArray(card.id, exportcards) === -1) {
                                                    console.log('skip card ' + card.id);
                                                    return true;
                                                }
                                            }

                                            var accept = true;
                                            if (filterListsNames.length > 0 && filterMode === 'Card') {
                                                for (var y = 0; y < filterListsNames.length; y++) {
                                                    if (!card.name.toLowerCase().startsWith(filterListsNames[y].trim().toLowerCase())) {
                                                        accept = false;
                                                    } else {
                                                        accept = true;
                                                        break;
                                                    }
                                                }
                                            }
                                            if (!accept) {
                                                console.log('skipping card ' + card.name);
                                                return true;
                                            }

                                            exportedCardsIDs.push(card.id);
                                            var title = card.name;

                                            console.log(i + ') Card #' + card.idShort + ' ' + title + ' (' + card.id + ') - POS: ' + card.pos);
                                            nProcessedCards++;

                                            var spent = 0;
                                            var estimate = 0;
                                            var checkListsText = '',
                                                commentsText = '',
                                                attachmentsText = '',
                                                memberCreator = '',
                                                datetimeCreated = null,
                                                memberDone = '',
                                                datetimeDone = '',
                                                jsonCheckLists = [],
                                                jsonComments = [],
                                                jsonAttachments = [];

                                            // Trello Plus Spent/Estimate
                                            spent = extractFloat(title, SpentEstRegex, 2);
                                            estimate = extractFloat(title, SpentEstRegex, 5);

                                            // Scrum for Trello Points Estimate/Consumed
                                            points_estimate = extractFloat(title, PointsEstRegex, 2);
                                            points_consumed = extractFloat(title, PointsConRegex, 2);

                                            // Clean-up title
                                            title = title.replace(SpentEstRegex, '');
                                            title = title.replace(PointsEstRegex, '');
                                            title = title.replace(PointsConRegex, '');
                                            title = title.trim();

                                            // tag archived cards
                                            if (card.closed) {
                                                title = '[archived] ' + title;
                                            }
                                            var due = card.due || '';

                                            //Get all the Member IDs
                                            // console.log('Members: ' + card.idMembers.length);
                                            var memberIDs = card.idMembers;
                                            var memberInitials = [];
                                            $.each(memberIDs, function(i, memberID) {
                                                $.each(card.members, function(key, member) {
                                                    if (member.id == memberID) {
                                                        if (member.fullName !== undefined) {
                                                            memberInitials.push(member.fullName); // initials, username or fullName    
                                                        } else {
                                                            if (member.username !== undefined) {
                                                                memberInitials.push(member.username);
                                                            }
                                                        }

                                                    }
                                                });
                                            });

                                            //Get all labels
                                            var labels = [];
                                            if (card.labels.length <= 0 && filterListsNames.length > 0 && filterMode === 'Label') {
                                                // filtering by label name: skip cards without labels
                                                accept = false;
                                                return true;
                                            }

                                            if (filterListsNames.length > 0 && filterMode === 'Label')
                                                accept = false;

                                            $.each(card.labels, function(i, label) {

                                                if (label.name) {
                                                    labels.push(label.name);
                                                } else {
                                                    labels.push(label.color);
                                                }

                                                if (filterListsNames.length > 0 && filterMode === 'Label') {
                                                    for (var y = 0; y < filterListsNames.length; y++) {
                                                        if (accept)
                                                            continue;
                                                        if (!label.name.toLowerCase().startsWith(filterListsNames[y].trim().toLowerCase())) {
                                                            accept = false;
                                                        } else {
                                                            accept = true;
                                                            break;
                                                        }
                                                    }
                                                }


                                            });
                                            if (!accept) {
                                                // filtering by label name
                                                // console.log('FILTER2 label filter skipping card ' + card.name);
                                                return true;
                                            }

                                            if (bExportChecklists) {
                                                //all checklists
                                                // console.log('Checklists: ' + card.idChecklists.length);
                                                var checklists = [];
                                                if (card.idChecklists !== undefined) {
                                                    $.each(card.idChecklists, function(i, idchecklist) {
                                                        if (idchecklist) {
                                                            checklists.push(idchecklist);
                                                        }
                                                    });
                                                }

                                                //parse checklists
                                                $.each(checklists, function(i, checklistid) {
                                                    // console.log('PARSE ' + checklistid);
                                                    $.each(card.checklists, function(key, list) {
                                                        var list_id = list.id;
                                                        //console.log('found ' + list_id);
                                                        if (list_id == checklistid) {
                                                            var jsonCheckList = {};
                                                            jsonCheckList.name = list.name;
                                                            jsonCheckList.items = [];
                                                            checkListsText += list.name + ' (' + list.checkItems.length + ' items):\n';
                                                            //checkitems: reordered (issue #4 https://github.com/trapias/trelloExport/issues/4)
                                                            var orderedChecklists = Enumerable.From(list.checkItems)
                                                                .OrderBy(function(x) {
                                                                    return x.pos;
                                                                })
                                                                .ToArray();

                                                            $.each(orderedChecklists, function(i, item) {
                                                                if (item) {
                                                                    var cItem = {};
                                                                    cItem.name = item.name;
                                                                    if (item.state == 'complete') {
                                                                        // issue #5
                                                                        // find who and when item was completed
                                                                        var oCompletedBy = searchupdateCheckItemStateOnCardAction(item.id, card.actions);
                                                                        checkListsText += ' - ' + item.name + ' [' + item.state + ' ' + oCompletedBy.date + ' by ' + oCompletedBy.by + ']\n';
                                                                        cItem.completed = true;
                                                                        cItem.completedDate = oCompletedBy.date;
                                                                        cItem.completedBy = oCompletedBy.by;
                                                                        jsonCheckList.items.push(cItem);

                                                                    } else {
                                                                        checkListsText += ' - ' + item.name + ' [' + item.state + ']\n';
                                                                        cItem.completed = false;
                                                                        jsonCheckList.items.push(cItem);
                                                                    }
                                                                }
                                                            });
                                                            jsonCheckLists.push(jsonCheckList);
                                                        }
                                                    });
                                                });
                                            }

                                            var numberOfComments = 0;

                                            if (bExportComments) {
                                                //comments
                                                var commentsOnCard = getCommentCardActions(idBoard, card.id);
                                                if (commentsOnCard) {
                                                    // console.log('parsing comments');
                                                    // console.log('parse ' + commentsOnCard.length + ' comment actions for this card ' + JSON.stringify(commentsOnCard));
                                                    $.each(commentsOnCard, function(j, action) {
                                                        if ((action.type == "commentCard" || action.type == 'copyCommentCard')) {
                                                            if (card.id == action.data.card.id) {
                                                                numberOfComments++;
                                                                var jsonComment = {};
                                                                //2013-08-08T06:57:18
                                                                var d = new Date(action.date);
                                                                jsonComment.date = d;
                                                                jsonComment.text = action.data.text;
                                                                var sActionDate = d.toLocaleDateString() + ' ' + d.toLocaleTimeString();
                                                                if (action.memberCreator !== undefined) {
                                                                    jsonComment.memberCreator = action.memberCreator;
                                                                    if (jsonComment.memberCreator.fullName === undefined) {
                                                                        jsonComment.memberCreator.fullName = jsonComment.memberCreator.username;
                                                                    }
                                                                    commentsText += '[' + sActionDate + ' - ' + (action.memberCreator.fullName !== undefined ? action.memberCreator.fullName : action.memberCreator.username) + '] ' + action.data.text + "\n";
                                                                } else {
                                                                    jsonComment.memberCreator = null;
                                                                    commentsText += '[' + sActionDate + '] ' + action.data.text + "\n";
                                                                }
                                                                jsonComments.push(jsonComment);
                                                            }
                                                        }
                                                    });
                                                }
                                            }

                                            if (bExportAttachments) {
                                                if (card.attachments) {
                                                    // console.log('Attachments: ' + card.attachments.length);
                                                    $.each(card.attachments, function(j, attach) {
                                                        // console.log(attach);
                                                        jsonAttachments.push(attach);
                                                        attachmentsText += '[' + attach.name + '] (' + attach.bytes + ') ' + attach.url + '\n';
                                                    });
                                                }
                                            }

                                            //pulled from https://github.com/bmccormack/export-for-trello/blob/5b2b8b102b98ed2c49241105cb9e00e44d4e1e86/trelloexport.js
                                            //Get member created and DateTime created
                                            var query = Enumerable.From(card.actions)
                                                .Where(function(x) {
                                                    if (x.data.card) {
                                                        return x.data.card.id == card.id && x.type == "createCard";
                                                    }
                                                })
                                                .ToArray();
                                            if (query.length > 0) {
                                                if (query[0].memberCreator !== undefined) {

                                                    if (query[0].memberCreator.fullName !== undefined) {
                                                        memberCreator = query[0].memberCreator.fullName + ' (' + query[0].memberCreator.username + ')';
                                                    } else {
                                                        memberCreator = query[0].memberCreator.username;
                                                    }
                                                    datetimeCreated = new Date(query[0].date);
                                                }

                                            } else {
                                                //use the API to get the action created method
                                                var actionCreateCard = getCreateCardAction(idBoard, card.id);
                                                if (actionCreateCard && actionCreateCard.memberCreator !== undefined) {
                                                    memberCreator = (actionCreateCard.memberCreator.fullName !== undefined ? actionCreateCard.memberCreator.fullName : actionCreateCard.memberCreator.username);
                                                    datetimeCreated = new Date(actionCreateCard.date);
                                                } else {
                                                    // calculate datetimeCreated from card id
                                                    // cfr http://help.trello.com/article/759-getting-the-time-a-card-or-board-was-created
                                                    datetimeCreated = new Date(1000 * parseInt(card.id.substring(0, 8), 16));
                                                    memberCreator = "";
                                                }
                                            }

                                            /**
                                             * 1.9.14: handle multiple nameListDone
                                             * e.g. Done, Finished
                                             */
                                            if (nameListDone === '') {
                                                nameListDone = 'Done';
                                            } // default
                                            var allnameListDone = nameListDone.split(',');
                                            for (var nd = 0; nd < allnameListDone.length; nd++) {
                                                // console.log(nd + ') ' + allnameListDone[nd]);
                                                if (memberDone === "") {
                                                    //Find out when the card was most recently moved to any list whose name starts with "Done" (ignore case, e.g. 'done' or 'DONE' or 'DoNe')
                                                    query = Enumerable.From(card.actions)
                                                        .Where(function(x) {
                                                            if (x.data.card && x.data.listAfter) {
                                                                var listAfterName = x.data.listAfter.name;
                                                                return x.data.card.id == card.id && listAfterName.toLowerCase().startsWith(allnameListDone[nd].trim().toLowerCase());
                                                            }
                                                        })
                                                        .OrderByDescending(function(x) {
                                                            return x.date;
                                                        })
                                                        .ToArray();
                                                    if (query.length > 0 && query[0].memberCreator !== undefined) {
                                                        memberDone = (query[0].memberCreator.fullName !== undefined ? query[0].memberCreator.fullName : query[0].memberCreator.username);
                                                        datetimeDone = query[0].date;
                                                    } else {
                                                        var actionMoveCard = getMoveCardAction(idBoard, card.id, allnameListDone[nd].trim());
                                                        if (actionMoveCard && actionMoveCard.memberCreator !== undefined) {
                                                            memberDone = (actionMoveCard.memberCreator.fullName !== undefined ? actionMoveCard.memberCreator.fullName : actionMoveCard.memberCreator.username);
                                                            datetimeDone = actionMoveCard.date;
                                                        } else {
                                                            memberDone = "";
                                                            datetimeDone = "";
                                                        }
                                                    }
                                                }
                                            }

                                            var completionTime = "",
                                                completionTimeText = "";
                                            if (datetimeDone !== "" && datetimeCreated !== "") {
                                                // var d1 = new Date(datetimeCreated);
                                                var d2 = new Date(datetimeDone);
                                                var df = new DateDiff(d2, datetimeCreated);
                                                // PnYnMnDTnHnMnS ISO8601 -> PnDTnHnMnS
                                                completionTime = "P" + df.days + "DT" + df.hours + "H" + df.minutes + "M" + df.seconds + "S";
                                                completionTimeText = df.days + ' days, ' + df.hours + ' hours, ' + df.minutes + ' minutes, ' + df.seconds + ' seconds';
                                                datetimeCreated = datetimeCreated.toLocaleDateString() + ' ' + datetimeCreated.toLocaleTimeString();
                                                datetimeDone = d2.toLocaleDateString() + ' ' + d2.toLocaleTimeString();
                                            } else {
                                                if (datetimeCreated) {
                                                    datetimeCreated = datetimeCreated.toLocaleDateString() + ' ' + datetimeCreated.toLocaleTimeString();
                                                }
                                            }

                                            var isArchived = false;
                                            if (list.closed || card.closed) {
                                                isArchived = true;
                                            }

                                            var dateLastActivity = new Date(card.dateLastActivity);

                                            var rowData = {
                                                'boardName': boardName,
                                                'listName': listName,
                                                'cardID': card.idShort,
                                                'title': title,
                                                'shortLink': 'https://trello.com/c/' + card.shortLink,
                                                'cardDescription': card.desc.substr(0, MAXCHARSPERCELL),
                                                'checkLists': checkListsText.substr(0, MAXCHARSPERCELL),
                                                'numberOfComments': numberOfComments,
                                                'comments': commentsText.substr(0, MAXCHARSPERCELL),
                                                'attachments': attachmentsText.substr(0, MAXCHARSPERCELL),
                                                'votes': card.idMembersVoted.length,
                                                'spent': spent,
                                                'estimate': estimate,
                                                'points_estimate': points_estimate,
                                                'points_consumed': points_consumed,
                                                'datetimeCreated': datetimeCreated,
                                                'memberCreator': memberCreator,
                                                'LastActivity': dateLastActivity.toLocaleDateString() + ' ' + dateLastActivity.toLocaleTimeString(),
                                                'due': due,
                                                'datetimeDone': datetimeDone,
                                                'memberDone': memberDone,
                                                'completionTime': completionTime,
                                                'completionTimeText': completionTimeText,
                                                'memberInitials': memberInitials.toString(),
                                                'labels': labels, //.toString(),
                                                'isArchived': isArchived,
                                                'jsonCheckLists': jsonCheckLists,
                                                'jsonComments': jsonComments,
                                                'jsonAttachments': jsonAttachments
                                            };

                                            jsonComputedCards.push(rowData);
                                        }
                                    });

                                })
                                .fail(function(jqXHR, textStatus, errorThrown) {
                                    console.log("Error!!!");
                                    readCards = -1;
                                    $.growl.error({
                                        title: "TrelloExport",
                                        message: jqXHR.statusText + ' ' + jqXHR.status + ': ' + jqXHR.responseText,
                                        fixed: true
                                    });
                                });

                        } while (readCards > 0);
                        // cards loop end

                    });

                })
                .fail(function(jqXHR, textStatus, errorThrown) {
                    console.log("Error!!!");
                    $.growl.error({
                        title: "TrelloExport",
                        message: jqXHR.statusText + ' ' + jqXHR.status + ': ' + jqXHR.responseText,
                        fixed: true
                    });
                });

            // end loop boards
            nProcessedBoards++;
        }

        console.log('Processed ' + nProcessedLists + ' lists and ' + nProcessedCards + ' cards');
        console.log("The End. Now export.");

        $.growl.notice({
            title: "TrelloExport",
            message: 'Done. Processed ' + nProcessedLists + ' lists and ' + nProcessedCards + ' cards in ' + nProcessedBoards + (nProcessedBoards > 1 ? ' boards.' : ' board.'),
            fixed: false
        });

        switch (exportFormat) {

            case 'XLSX':
                createExcelExport(jsonComputedCards, iExcelItemsAsRows, allColumns, selectedColumns);
                break;

            case 'MD':
                createMarkdownExport(jsonComputedCards, true, true, bchkHTMLInlineImages);
                break;

            case 'HTML':
                createHTMLExport(jsonComputedCards, bckHTMLCardInfo, bchkHTMLInlineImages, css);
                break;

            case 'OPML':
                createOPMLExport(jsonComputedCards);
                break;

            default:
                console.log('Unknown exportFormat requested');
                $.growl.error({
                    title: "TrelloExport",
                    message: 'Unknown exportFormat requested',
                    fixed: true
                });
                break;
        }


    }).catch(
        // rejection
        function(reason) {
            console.log('===Error=== ' + reason);
            $.growl.error({
                title: "TrelloExport",
                message: reason,
                fixed: true
            });
        });

}

// createExcelExport: export to XLSX
function createExcelExport(jsonComputedCards, iExcelItemsAsRows, allColumns, columnHeadings) {
    console.log('TrelloExport exporting to Excel ' + jsonComputedCards.length + ' cards...');

    // prepare Workbook
    var wb = new Workbook();
    wArchived = {};
    wArchived.name = 'Archived lists and cards';
    wArchived.data = [];
    wArchived.data.push([]);
    wArchived.data[0] = columnHeadings;

    // Setup the active list and cart worksheet
    w = {};
    // if(data.name.length>30)
    //     w.name = data.name.substr(0,30);
    // else
    //     w.name = data.name;
    w.data = [];
    w.data.push([]);
    w.data[0] = columnHeadings;

    // loop jsonComputedCards
    jsonComputedCards.forEach(function(card) {

        var toStringArray = [];
        var nTotalCheckListItems = 0,
            nTotalCheckListItemsCompleted = 0;

        card.jsonCheckLists.forEach(function(list) {
            nTotalCheckListItems += list.items.length;

            list.items.forEach(function(it) {
                if (it.completed) {
                    nTotalCheckListItemsCompleted++;
                }
            });
        });

        // iExcelItemsAsRows: 
        // 0=default
        // 1=checklist items as rows
        // 2=labels as rows
        // 3=members as rows
        switch (Number(iExcelItemsAsRows)) {
            case 0:
                // standard checklist export

                // toStringArray = [
                //     card.boardName,
                //     card.listName,
                //     card.cardID,
                //     card.title,
                //     card.shortLink,
                //     card.cardDescription,
                //     nTotalCheckListItems,
                //     nTotalCheckListItemsCompleted,
                //     card.checkLists,
                //     card.numberOfComments,
                //     card.comments,
                //     card.attachments,
                //     card.votes,
                //     card.spent,
                //     card.estimate,
                //     card.points_estimate,
                //     card.points_consumed,
                //     card.datetimeCreated,
                //     card.memberCreator,
                //     card.LastActivity,
                //     (card.due !== '' ? new Date(card.due).toLocaleDateString() + ' ' + new Date(card.due).toLocaleTimeString() : ''),
                //     card.datetimeDone,
                //     card.memberDone,
                //     card.completionTime,
                //     card.memberInitials,
                //     card.labels.toString()
                //     // ,card.isArchived
                // ];

                toStringArray = [];
                // filter columns
                for (var nCol = 0; nCol < allColumns.length; nCol++) {
                    if ($.inArray(allColumns[nCol].value, columnHeadings) > -1) {

                        switch (nCol) {
                            case 0:
                                toStringArray.push(card.boardName);
                                break;
                            case 1:
                                toStringArray.push(card.listName);
                                break;
                            case 2:
                                toStringArray.push(card.cardID);
                                break;
                            case 3:
                                toStringArray.push(card.title);
                                break;
                            case 4:
                                toStringArray.push(card.shortLink);
                                break;
                            case 5:
                                toStringArray.push(card.cardDescription);
                                break;
                            case 6:
                                toStringArray.push(card.nTotalCheckListItems);
                                break;
                            case 7:
                                toStringArray.push(card.nTotalCheckListItemsCompleted);
                                break;
                            case 8:
                                toStringArray.push(card.checkLists);
                                break;
                            case 9:
                                toStringArray.push(card.numberOfComments);
                                break;
                            case 10:
                                toStringArray.push(card.comments);
                                break;
                            case 11:
                                toStringArray.push(card.attachments);
                                break;
                            case 12:
                                toStringArray.push(card.votes);
                                break;
                            case 13:
                                toStringArray.push(card.spent);
                                break;
                            case 14:
                                toStringArray.push(card.estimate);
                                break;
                            case 15:
                                toStringArray.push(card.points_estimate);
                                break;
                            case 16:
                                toStringArray.push(card.points_consumed);
                                break;
                            case 17:
                                toStringArray.push(card.datetimeCreated);
                                break;
                            case 18:
                                toStringArray.push(card.memberCreator);
                                break;
                            case 19:
                                toStringArray.push(card.LastActivity);
                                break;
                            case 20:
                                toStringArray.push((card.due !== '' ? new Date(card.due).toLocaleDateString() + ' ' + new Date(card.due).toLocaleTimeString() : ''));
                                break;
                            case 21:
                                toStringArray.push(card.datetimeDone);
                                break;
                            case 22:
                                toStringArray.push(card.memberDone);
                                break;
                            case 23:
                                toStringArray.push(card.completionTime);
                                break;
                            case 24:
                                toStringArray.push(card.memberInitials);
                                break;
                            case 25:
                                toStringArray.push(card.labels.toString());
                                break;
                        }
                    }
                }

                if (card.isArchived) {
                    var rArch = wArchived.data.push([]) - 1;
                    wArchived.data[rArch] = toStringArray;
                } else {
                    var r = w.data.push([]) - 1;
                    w.data[r] = toStringArray;
                }
                break;

            case 1:
                // checklist items as rows
                if (card.jsonCheckLists.length > 0) {

                    card.jsonCheckLists.forEach(function(list) {
                        list.items.forEach(function(it) {

                            // toStringArray = [
                            //     card.boardName,
                            //     card.listName,
                            //     card.cardID,
                            //     card.title,
                            //     card.shortLink,
                            //     card.cardDescription,
                            //     nTotalCheckListItems,
                            //     nTotalCheckListItemsCompleted,
                            //     list.name,
                            //     it.name,
                            //     it.completed,
                            //     it.completedDate,
                            //     it.completedBy,
                            //     card.numberOfComments,
                            //     card.comments,
                            //     card.attachments,
                            //     card.votes,
                            //     card.spent,
                            //     card.estimate,
                            //     card.points_estimate,
                            //     card.points_consumed,
                            //     card.datetimeCreated,
                            //     card.memberCreator,
                            //     card.LastActivity,
                            //     card.due,
                            //     card.datetimeDone,
                            //     card.memberDone,
                            //     card.completionTime,
                            //     card.memberInitials,
                            //     card.labels.toString()
                            //     // ,card.isArchived
                            // ];

                            toStringArray = [];
                            // filter columns
                            for (var nCol = 0; nCol < allColumns.length; nCol++) {
                                if ($.inArray(allColumns[nCol].value, columnHeadings) > -1) {

                                    switch (nCol) {
                                        case 0:
                                            toStringArray.push(card.boardName);
                                            break;
                                        case 1:
                                            toStringArray.push(card.listName);
                                            break;
                                        case 2:
                                            toStringArray.push(card.cardID);
                                            break;
                                        case 3:
                                            toStringArray.push(card.title);
                                            break;
                                        case 4:
                                            toStringArray.push(card.shortLink);
                                            break;
                                        case 5:
                                            toStringArray.push(card.cardDescription);
                                            break;
                                        case 6:
                                            toStringArray.push(card.nTotalCheckListItems);
                                            break;
                                        case 7:
                                            toStringArray.push(card.nTotalCheckListItemsCompleted);
                                            break;
                                        case 8:
                                            toStringArray.push(list.name);
                                            break;
                                        case 9:
                                            toStringArray.push(it.name);
                                            break;
                                        case 10:
                                            toStringArray.push(it.completed);
                                            break;
                                        case 11:
                                            toStringArray.push(it.completedDate);
                                            break;
                                        case 12:
                                            toStringArray.push(it.completedBy);
                                            break;
                                        case 13:
                                            toStringArray.push(card.numberOfComments);
                                            break;
                                        case 14:
                                            toStringArray.push(card.comments);
                                            break;
                                        case 15:
                                            toStringArray.push(card.attachments);
                                            break;
                                        case 16:
                                            toStringArray.push(card.votes);
                                            break;
                                        case 17:
                                            toStringArray.push(card.spent);
                                            break;
                                        case 18:
                                            toStringArray.push(card.estimate);
                                            break;
                                        case 19:
                                            toStringArray.push(card.points_estimate);
                                            break;
                                        case 20:
                                            toStringArray.push(card.points_consumed);
                                            break;
                                        case 21:
                                            toStringArray.push(card.datetimeCreated);
                                            break;
                                        case 22:
                                            toStringArray.push(card.memberCreator);
                                            break;
                                        case 23:
                                            toStringArray.push(card.LastActivity);
                                            break;
                                        case 24:
                                            toStringArray.push((card.due !== '' ? new Date(card.due).toLocaleDateString() + ' ' + new Date(card.due).toLocaleTimeString() : ''));
                                            break;
                                        case 25:
                                            toStringArray.push(card.datetimeDone);
                                            break;
                                        case 26:
                                            toStringArray.push(card.memberDone);
                                            break;
                                        case 27:
                                            toStringArray.push(card.completionTime);
                                            break;
                                        case 28:
                                            toStringArray.push(card.memberInitials);
                                            break;
                                        case 29:
                                            toStringArray.push(card.labels.toString());
                                            break;
                                    }
                                }
                            }

                            if (card.isArchived) {
                                var rArch2 = wArchived.data.push([]) - 1;
                                wArchived.data[rArch2] = toStringArray;
                            } else {
                                var r2 = w.data.push([]) - 1;
                                w.data[r2] = toStringArray;
                            }

                        });
                    });

                } else {
                    // no checklist items
                    // toStringArray = [
                    //     card.boardName,
                    //     card.listName,
                    //     card.cardID,
                    //     card.title,
                    //     card.shortLink,
                    //     card.cardDescription,
                    //     nTotalCheckListItems,
                    //     nTotalCheckListItemsCompleted,
                    //     '',
                    //     '',
                    //     '',
                    //     '',
                    //     '',
                    //     card.numberOfComments,
                    //     card.comments,
                    //     card.attachments,
                    //     card.votes,
                    //     card.spent,
                    //     card.estimate,
                    //     card.points_estimate,
                    //     card.points_consumed,
                    //     card.datetimeCreated,
                    //     card.memberCreator,
                    //     card.LastActivity,
                    //     card.due,
                    //     card.datetimeDone,
                    //     card.memberDone,
                    //     card.completionTime,
                    //     card.memberInitials,
                    //     card.labels.toString()
                    //     // ,card.isArchived
                    // ];

                    toStringArray = [];
                    // filter columns
                    for (nCol = 0; nCol < allColumns.length; nCol++) {
                        if ($.inArray(allColumns[nCol].value, columnHeadings) > -1) {

                            switch (nCol) {
                                case 0:
                                    toStringArray.push(card.boardName);
                                    break;
                                case 1:
                                    toStringArray.push(card.listName);
                                    break;
                                case 2:
                                    toStringArray.push(card.cardID);
                                    break;
                                case 3:
                                    toStringArray.push(card.title);
                                    break;
                                case 4:
                                    toStringArray.push(card.shortLink);
                                    break;
                                case 5:
                                    toStringArray.push(card.cardDescription);
                                    break;
                                case 6:
                                    toStringArray.push(card.nTotalCheckListItems);
                                    break;
                                case 7:
                                    toStringArray.push(card.nTotalCheckListItemsCompleted);
                                    break;
                                case 8:
                                    toStringArray.push('');
                                    break;
                                case 9:
                                    toStringArray.push('');
                                    break;
                                case 10:
                                    toStringArray.push('');
                                    break;
                                case 11:
                                    toStringArray.push('');
                                    break;
                                case 12:
                                    toStringArray.push('');
                                    break;
                                case 13:
                                    toStringArray.push(card.numberOfComments);
                                    break;
                                case 14:
                                    toStringArray.push(card.comments);
                                    break;
                                case 15:
                                    toStringArray.push(card.attachments);
                                    break;
                                case 16:
                                    toStringArray.push(card.votes);
                                    break;
                                case 17:
                                    toStringArray.push(card.spent);
                                    break;
                                case 18:
                                    toStringArray.push(card.estimate);
                                    break;
                                case 19:
                                    toStringArray.push(card.points_estimate);
                                    break;
                                case 20:
                                    toStringArray.push(card.points_consumed);
                                    break;
                                case 21:
                                    toStringArray.push(card.datetimeCreated);
                                    break;
                                case 22:
                                    toStringArray.push(card.memberCreator);
                                    break;
                                case 23:
                                    toStringArray.push(card.LastActivity);
                                    break;
                                case 24:
                                    toStringArray.push((card.due !== '' ? new Date(card.due).toLocaleDateString() + ' ' + new Date(card.due).toLocaleTimeString() : ''));
                                    break;
                                case 25:
                                    toStringArray.push(card.datetimeDone);
                                    break;
                                case 26:
                                    toStringArray.push(card.memberDone);
                                    break;
                                case 27:
                                    toStringArray.push(card.completionTime);
                                    break;
                                case 28:
                                    toStringArray.push(card.memberInitials);
                                    break;
                                case 29:
                                    toStringArray.push(card.labels.toString());
                                    break;
                            }
                        }
                    }

                    if (card.isArchived) {
                        var rArch2 = wArchived.data.push([]) - 1;
                        wArchived.data[rArch2] = toStringArray;
                    } else {
                        var r2 = w.data.push([]) - 1;
                        w.data[r2] = toStringArray;
                    }
                }
                break;

            case 2:
                // 2=labels as rows
                if (card.labels.length > 0) {

                    card.labels.forEach(function(lbl) {

                        // toStringArray = [
                        //     card.boardName,
                        //     card.listName,
                        //     card.cardID,
                        //     card.title,
                        //     card.shortLink,
                        //     card.cardDescription,
                        //     nTotalCheckListItems,
                        //     nTotalCheckListItemsCompleted,
                        //     card.checkLists,
                        //     card.numberOfComments,
                        //     card.comments,
                        //     card.attachments,
                        //     card.votes,
                        //     card.spent,
                        //     card.estimate,
                        //     card.points_estimate,
                        //     card.points_consumed,
                        //     card.datetimeCreated,
                        //     card.memberCreator,
                        //     card.LastActivity,
                        //     (card.due !== '' ? new Date(card.due).toLocaleDateString() + ' ' + new Date(card.due).toLocaleTimeString() : ''),
                        //     card.datetimeDone,
                        //     card.memberDone,
                        //     card.completionTime,
                        //     card.memberInitials,
                        //     lbl
                        // ];

                        toStringArray = [];
                        // filter columns
                        for (var nCol = 0; nCol < allColumns.length; nCol++) {
                            if ($.inArray(allColumns[nCol].value, columnHeadings) > -1) {

                                switch (nCol) {
                                    case 0:
                                        toStringArray.push(card.boardName);
                                        break;
                                    case 1:
                                        toStringArray.push(card.listName);
                                        break;
                                    case 2:
                                        toStringArray.push(card.cardID);
                                        break;
                                    case 3:
                                        toStringArray.push(card.title);
                                        break;
                                    case 4:
                                        toStringArray.push(card.shortLink);
                                        break;
                                    case 5:
                                        toStringArray.push(card.cardDescription);
                                        break;
                                    case 6:
                                        toStringArray.push(card.nTotalCheckListItems);
                                        break;
                                    case 7:
                                        toStringArray.push(card.nTotalCheckListItemsCompleted);
                                        break;
                                    case 8:
                                        toStringArray.push(card.checkLists);
                                        break;
                                    case 9:
                                        toStringArray.push(card.numberOfComments);
                                        break;
                                    case 10:
                                        toStringArray.push(card.comments);
                                        break;
                                    case 11:
                                        toStringArray.push(card.attachments);
                                        break;
                                    case 12:
                                        toStringArray.push(card.votes);
                                        break;
                                    case 13:
                                        toStringArray.push(card.spent);
                                        break;
                                    case 14:
                                        toStringArray.push(card.estimate);
                                        break;
                                    case 15:
                                        toStringArray.push(card.points_estimate);
                                        break;
                                    case 16:
                                        toStringArray.push(card.points_consumed);
                                        break;
                                    case 17:
                                        toStringArray.push(card.datetimeCreated);
                                        break;
                                    case 18:
                                        toStringArray.push(card.memberCreator);
                                        break;
                                    case 19:
                                        toStringArray.push(card.LastActivity);
                                        break;
                                    case 20:
                                        toStringArray.push((card.due !== '' ? new Date(card.due).toLocaleDateString() + ' ' + new Date(card.due).toLocaleTimeString() : ''));
                                        break;
                                    case 21:
                                        toStringArray.push(card.datetimeDone);
                                        break;
                                    case 22:
                                        toStringArray.push(card.memberDone);
                                        break;
                                    case 23:
                                        toStringArray.push(card.completionTime);
                                        break;
                                    case 24:
                                        toStringArray.push(card.memberInitials);
                                        break;
                                    case 25:
                                        toStringArray.push(lbl);
                                        break;
                                }
                            }
                        }

                        if (card.isArchived) {
                            var rArch2 = wArchived.data.push([]) - 1;
                            wArchived.data[rArch2] = toStringArray;
                        } else {
                            var r2 = w.data.push([]) - 1;
                            w.data[r2] = toStringArray;
                        }

                    });

                } else {
                    // no labels
                    // toStringArray = [
                    //     card.boardName,
                    //     card.listName,
                    //     card.cardID,
                    //     card.title,
                    //     card.shortLink,
                    //     card.cardDescription,
                    //     nTotalCheckListItems,
                    //     nTotalCheckListItemsCompleted,
                    //     '',
                    //     '',
                    //     '',
                    //     '',
                    //     '',
                    //     card.comments,
                    //     card.numberOfComments,
                    //     card.attachments,
                    //     card.votes,
                    //     card.spent,
                    //     card.estimate,
                    //     card.points_estimate,
                    //     card.points_consumed,
                    //     card.datetimeCreated,
                    //     card.memberCreator,
                    //     card.LastActivity,
                    //     card.due,
                    //     card.datetimeDone,
                    //     card.memberDone,
                    //     card.completionTime,
                    //     card.memberInitials,
                    //     card.labels.toString()
                    // ];

                    toStringArray = [];
                    // filter columns
                    for (nCol = 0; nCol < allColumns.length; nCol++) {
                        if ($.inArray(allColumns[nCol].value, columnHeadings) > -1) {

                            switch (nCol) {
                                case 0:
                                    toStringArray.push(card.boardName);
                                    break;
                                case 1:
                                    toStringArray.push(card.listName);
                                    break;
                                case 2:
                                    toStringArray.push(card.cardID);
                                    break;
                                case 3:
                                    toStringArray.push(card.title);
                                    break;
                                case 4:
                                    toStringArray.push(card.shortLink);
                                    break;
                                case 5:
                                    toStringArray.push(card.cardDescription);
                                    break;
                                case 6:
                                    toStringArray.push(card.nTotalCheckListItems);
                                    break;
                                case 7:
                                    toStringArray.push(card.nTotalCheckListItemsCompleted);
                                    break;
                                case 8:
                                    toStringArray.push('');
                                    break;
                                case 9:
                                    toStringArray.push('');
                                    break;
                                case 10:
                                    toStringArray.push('');
                                    break;
                                case 11:
                                    toStringArray.push('');
                                    break;
                                case 12:
                                    toStringArray.push('');
                                    break;
                                case 13:
                                    toStringArray.push(card.numberOfComments);
                                    break;
                                case 14:
                                    toStringArray.push(card.comments);
                                    break;
                                case 15:
                                    toStringArray.push(card.attachments);
                                    break;
                                case 16:
                                    toStringArray.push(card.votes);
                                    break;
                                case 17:
                                    toStringArray.push(card.spent);
                                    break;
                                case 18:
                                    toStringArray.push(card.estimate);
                                    break;
                                case 19:
                                    toStringArray.push(card.points_estimate);
                                    break;
                                case 20:
                                    toStringArray.push(card.points_consumed);
                                    break;
                                case 21:
                                    toStringArray.push(card.datetimeCreated);
                                    break;
                                case 22:
                                    toStringArray.push(card.memberCreator);
                                    break;
                                case 23:
                                    toStringArray.push(card.LastActivity);
                                    break;
                                case 24:
                                    toStringArray.push((card.due !== '' ? new Date(card.due).toLocaleDateString() + ' ' + new Date(card.due).toLocaleTimeString() : ''));
                                    break;
                                case 25:
                                    toStringArray.push(card.datetimeDone);
                                    break;
                                case 26:
                                    toStringArray.push(card.memberDone);
                                    break;
                                case 27:
                                    toStringArray.push(card.completionTime);
                                    break;
                                case 28:
                                    toStringArray.push(card.memberInitials);
                                    break;
                                case 29:
                                    toStringArray.push(card.labels.toString());
                                    break;
                            }
                        }
                    }

                    if (card.isArchived) {
                        var lblrArch2 = wArchived.data.push([]) - 1;
                        wArchived.data[lblrArch2] = toStringArray;
                    } else {
                        var lblr2 = w.data.push([]) - 1;
                        w.data[lblr2] = toStringArray;
                    }
                }
                break;

            case 3:
                // 3=members as rows
                if (card.memberInitials.length > 0) {

                    card.memberInitials.split(",").forEach(function(mbm) {

                        // toStringArray = [
                        //     card.boardName,
                        //     card.listName,
                        //     card.cardID,
                        //     card.title,
                        //     card.shortLink,
                        //     card.cardDescription,
                        //     nTotalCheckListItems,
                        //     nTotalCheckListItemsCompleted,
                        //     card.checkLists,
                        //     card.numberOfComments,
                        //     card.comments,
                        //     card.attachments,
                        //     card.votes,
                        //     card.spent,
                        //     card.estimate,
                        //     card.points_estimate,
                        //     card.points_consumed,
                        //     card.datetimeCreated,
                        //     card.memberCreator,
                        //     card.LastActivity,
                        //     (card.due !== '' ? new Date(card.due).toLocaleDateString() + ' ' + new Date(card.due).toLocaleTimeString() : ''),
                        //     card.datetimeDone,
                        //     card.memberDone,
                        //     card.completionTime,
                        //     mbm,
                        //     card.labels.toString()
                        // ];

                        toStringArray = [];
                        // filter columns
                        for (var nCol = 0; nCol < allColumns.length; nCol++) {
                            if ($.inArray(allColumns[nCol].value, columnHeadings) > -1) {

                                switch (nCol) {
                                    case 0:
                                        toStringArray.push(card.boardName);
                                        break;
                                    case 1:
                                        toStringArray.push(card.listName);
                                        break;
                                    case 2:
                                        toStringArray.push(card.cardID);
                                        break;
                                    case 3:
                                        toStringArray.push(card.title);
                                        break;
                                    case 4:
                                        toStringArray.push(card.shortLink);
                                        break;
                                    case 5:
                                        toStringArray.push(card.cardDescription);
                                        break;
                                    case 6:
                                        toStringArray.push(card.nTotalCheckListItems);
                                        break;
                                    case 7:
                                        toStringArray.push(card.nTotalCheckListItemsCompleted);
                                        break;
                                    case 8:
                                        toStringArray.push(card.checkLists);
                                        break;
                                    case 9:
                                        toStringArray.push(card.numberOfComments);
                                        break;
                                    case 10:
                                        toStringArray.push(card.comments);
                                        break;
                                    case 11:
                                        toStringArray.push(card.attachments);
                                        break;
                                    case 12:
                                        toStringArray.push(card.votes);
                                        break;
                                    case 13:
                                        toStringArray.push(card.spent);
                                        break;
                                    case 14:
                                        toStringArray.push(card.estimate);
                                        break;
                                    case 15:
                                        toStringArray.push(card.points_estimate);
                                        break;
                                    case 16:
                                        toStringArray.push(card.points_consumed);
                                        break;
                                    case 17:
                                        toStringArray.push(card.datetimeCreated);
                                        break;
                                    case 18:
                                        toStringArray.push(card.memberCreator);
                                        break;
                                    case 19:
                                        toStringArray.push(card.LastActivity);
                                        break;
                                    case 20:
                                        toStringArray.push((card.due !== '' ? new Date(card.due).toLocaleDateString() + ' ' + new Date(card.due).toLocaleTimeString() : ''));
                                        break;
                                    case 21:
                                        toStringArray.push(card.datetimeDone);
                                        break;
                                    case 22:
                                        toStringArray.push(card.memberDone);
                                        break;
                                    case 23:
                                        toStringArray.push(card.completionTime);
                                        break;
                                    case 24:
                                        toStringArray.push(mbm);
                                        break;
                                    case 25:
                                        toStringArray.push(card.labels.toString());
                                        break;
                                }
                            }
                        }

                        if (card.isArchived) {
                            var rArch2 = wArchived.data.push([]) - 1;
                            wArchived.data[rArch2] = toStringArray;
                        } else {
                            var r2 = w.data.push([]) - 1;
                            w.data[r2] = toStringArray;
                        }

                    });

                } else {
                    // no members
                    // toStringArray = [
                    //     card.boardName,
                    //     card.listName,
                    //     card.cardID,
                    //     card.title,
                    //     card.shortLink,
                    //     card.cardDescription,
                    //     nTotalCheckListItems,
                    //     nTotalCheckListItemsCompleted,
                    //     '',
                    //     '',
                    //     '',
                    //     '',
                    //     '',
                    //     card.numberOfComments,
                    //     card.comments,
                    //     card.attachments,
                    //     card.votes,
                    //     card.spent,
                    //     card.estimate,
                    //     card.points_estimate,
                    //     card.points_consumed,
                    //     card.datetimeCreated,
                    //     card.memberCreator,
                    //     card.LastActivity,
                    //     card.due,
                    //     card.datetimeDone,
                    //     card.memberDone,
                    //     card.completionTime,
                    //     card.memberInitials,
                    //     card.labels.toString()
                    // ];

                    toStringArray = [];
                    // filter columns
                    for (nCol = 0; nCol < allColumns.length; nCol++) {
                        if ($.inArray(allColumns[nCol].value, columnHeadings) > -1) {

                            switch (nCol) {
                                case 0:
                                    toStringArray.push(card.boardName);
                                    break;
                                case 1:
                                    toStringArray.push(card.listName);
                                    break;
                                case 2:
                                    toStringArray.push(card.cardID);
                                    break;
                                case 3:
                                    toStringArray.push(card.title);
                                    break;
                                case 4:
                                    toStringArray.push(card.shortLink);
                                    break;
                                case 5:
                                    toStringArray.push(card.cardDescription);
                                    break;
                                case 6:
                                    toStringArray.push(card.nTotalCheckListItems);
                                    break;
                                case 7:
                                    toStringArray.push(card.nTotalCheckListItemsCompleted);
                                    break;
                                case 8:
                                    toStringArray.push('');
                                    break;
                                case 9:
                                    toStringArray.push('');
                                    break;
                                case 10:
                                    toStringArray.push('');
                                    break;
                                case 11:
                                    toStringArray.push('');
                                    break;
                                case 12:
                                    toStringArray.push('');
                                    break;
                                case 13:
                                    toStringArray.push(card.numberOfComments);
                                    break;
                                case 14:
                                    toStringArray.push(card.comments);
                                    break;
                                case 15:
                                    toStringArray.push(card.attachments);
                                    break;
                                case 16:
                                    toStringArray.push(card.votes);
                                    break;
                                case 17:
                                    toStringArray.push(card.spent);
                                    break;
                                case 18:
                                    toStringArray.push(card.estimate);
                                    break;
                                case 19:
                                    toStringArray.push(card.points_estimate);
                                    break;
                                case 20:
                                    toStringArray.push(card.points_consumed);
                                    break;
                                case 21:
                                    toStringArray.push(card.datetimeCreated);
                                    break;
                                case 22:
                                    toStringArray.push(card.memberCreator);
                                    break;
                                case 23:
                                    toStringArray.push(card.LastActivity);
                                    break;
                                case 24:
                                    toStringArray.push((card.due !== '' ? new Date(card.due).toLocaleDateString() + ' ' + new Date(card.due).toLocaleTimeString() : ''));
                                    break;
                                case 25:
                                    toStringArray.push(card.datetimeDone);
                                    break;
                                case 26:
                                    toStringArray.push(card.memberDone);
                                    break;
                                case 27:
                                    toStringArray.push(card.completionTime);
                                    break;
                                case 28:
                                    toStringArray.push(card.memberInitials);
                                    break;
                                case 29:
                                    toStringArray.push(card.labels.toString());
                                    break;
                            }
                        }
                    }

                    if (card.isArchived) {
                        var lblrArch3 = wArchived.data.push([]) - 1;
                        wArchived.data[lblrArch3] = toStringArray;
                    } else {
                        var lblr3 = w.data.push([]) - 1;
                        w.data[lblr3] = toStringArray;
                    }
                }
                break;

            default:
                // console.log('DEFAULT: ' + iExcelItemsAsRows);
                break;
        }

    });

    var board_title = "TrelloExport";
    var ws = sheet_from_array_of_arrays(w.data);

    // add worksheet to workbook
    wb.SheetNames.push(board_title);
    wb.Sheets[board_title] = ws;
    console.log("Added sheet " + board_title);

    //add the Archived data
    var wsArchived = sheet_from_array_of_arrays(wArchived.data);
    if (wsArchived !== undefined) {
        wb.SheetNames.push("Archive");
        wb.Sheets.Archive = wsArchived;
    }

    var now = new Date();
    var fileName = "TrelloExport_" + now.getFullYear() + dd(now.getMonth() + 1) + dd(now.getUTCDate()) + dd(now.getHours()) + dd(now.getMinutes()) + dd(now.getSeconds()) + ".xlsx";
    var wbout = XLSX.write(wb, {
        bookType: 'xlsx',
        bookSST: true,
        type: 'binary'
    });
    var blob = new Blob([s2ab(wbout)], {
        type: "application/octet-stream"
    });
    console.log('saving BLOB ' + blob);
    saveAs(blob, fileName);

    console.log('Done exporting ' + fileName);

    $.growl.notice({
        title: "TrelloExport",
        message: 'Done. Downloading xlsx file...',
        fixed: true
    });
}

function createMarkdownExport(jsonComputedCards, bPrint, bckHTMLCardInfo, bchkHTMLInlineImages) {
    console.log('TrelloExport exporting to Markdown ' + jsonComputedCards.length + ' cards...');

    // console.log('jsonComputedCards: ' + JSON.stringify(jsonComputedCards));
    var mdOut = '',
        prevBoard = null,
        prevList = null;

    // loop jsonComputedCards
    jsonComputedCards.forEach(function(card) {

        if (prevBoard !== card.boardName) {
            // print board title
            mdOut += '# ' + card.boardName.trim() + '\n\n';
            prevBoard = card.boardName;
        }

        if (prevList !== card.listName) {
            // print board title
            mdOut += '## ' + card.listName.trim() + '\n\n';
            prevList = card.listName;
        }

        var sTitle = '';
        if (!bPrint) {
            // HTML export: add link to card
            sTitle = '### [ <a target="_blank" href="' + card.shortLink + '">' + card.cardID + '</a> ] ' + card.title.trim() + '\n\n';
        } else {
            sTitle = '### [' + card.cardID + '] ' + card.title.trim() + '\n\n';
        }
        if (bckHTMLCardInfo) {
            sTitle += '**Created:** ' + card.datetimeCreated + '\n\n' +
                '**Created by:** ' + card.memberCreator + '\n\n';
        }

        mdOut += sTitle +
            (card.due !== '' ? '**Due:** ' + new Date(card.due).toLocaleDateString() + ' ' + new Date(card.due).toLocaleTimeString() : '') + '\n\n' +
            (card.datetimeDone !== '' ? '**Completed:** ' + card.datetimeDone + '\n\n' : '') +
            (card.memberDone !== '' ? '**Completed by:** ' + card.memberDone + '\n\n' : '') +
            (card.completionTimeText !== '' ? '**Elapse:** ' + card.completionTimeText + '\n\n' : '') +
            (card.cardDescription !== '' ? '**Description:**\n\n' + card.cardDescription + '\n\n' : '');

        // checklists
        var i;
        if (card.jsonCheckLists.length > 0) {
            var prevCL = null;
            for (i = 0; i < card.jsonCheckLists.length; i++) {

                if (prevCL !== card.jsonCheckLists[i].name) {
                    mdOut += '#### ' + card.jsonCheckLists[i].name + '\n\n';
                    prevCL = card.jsonCheckLists[i].name;
                }

                for (var ii = 0; ii < card.jsonCheckLists[i].items.length; ii++) {
                    if (card.jsonCheckLists[i].items[ii].completed === true) {
                        mdOut += '[x] ' + card.jsonCheckLists[i].items[ii].name + ' ' + card.jsonCheckLists[i].items[ii].completedDate + ' ' + card.jsonCheckLists[i].items[ii].completedBy + '\n\n';
                    } else {
                        mdOut += '[ ] ' + card.jsonCheckLists[i].items[ii].name + '\n\n';
                    }
                }
            }
        }

        // comments
        if (card.jsonComments.length > 0) {
            mdOut += '#### Comments\n';
            for (i = 0; i < card.jsonComments.length; i++) {
                var d = card.jsonComments[i].date;
                mdOut += '**' + d.toLocaleDateString() + ' ' + d.toLocaleTimeString() + ' ' + card.jsonComments[i].memberCreator.fullName + '**\n\n' + card.jsonComments[i].text + '\n\n';
            }
        }

        // attachments
        if (card.jsonAttachments.length > 0) {
            mdOut += '#### Attachments\n';
            for (i = 0; i < card.jsonAttachments.length; i++) {

                // console.log('ATTACHMENT = ' + JSON.stringify(card.jsonAttachments[i]) );
                if (bchkHTMLInlineImages && isImage(card.jsonAttachments[i].name)) {

                    var sImg = '![' + card.jsonAttachments[i].name + '](' + card.jsonAttachments[i].url + ')';

                    mdOut += '**' + card.jsonAttachments[i].name + '** (' + Number(card.jsonAttachments[i].bytes / 1024).toFixed(2) + ' kb) [download](' + card.jsonAttachments[i].url + ')\n\n' + sImg + '\n\n';

                } else {

                    mdOut += '**' + card.jsonAttachments[i].name + '** (' + Number(card.jsonAttachments[i].bytes / 1024).toFixed(2) + ' kb) [download](' + card.jsonAttachments[i].url + ')\n\n';

                }

            }
        }

    });

    if (!bPrint) {
        return mdOut;
    }

    var now = new Date();
    var fileName = "TrelloExport_" + now.getFullYear() + dd(now.getMonth() + 1) + dd(now.getUTCDate()) + dd(now.getHours()) + dd(now.getMinutes()) + dd(now.getSeconds()) + ".md";

    saveAs(new Blob([mdOut], {
        type: "text/markdown;charset=utf-8"
    }), fileName);

    console.log('Done exporting ' + fileName);

    $.growl.notice({
        title: "TrelloExport",
        message: 'Done. Downloading markdown file...',
        fixed: true
    });
}

function isImage(name) {
    name = name.toLowerCase();
    return (name.endsWith("jpg") || name.endsWith("jpeg") || name.endsWith("png"));
}

function createHTMLExport(jsonComputedCards, bckHTMLCardInfo, bchkHTMLInlineImages, css) {
    var md = createMarkdownExport(jsonComputedCards, false, bckHTMLCardInfo, bchkHTMLInlineImages);
    var converter = new showdown.Converter();
    html = converter.makeHtml(html_encode(md));

    if (css === undefined || css === '' || css === null) {
        css = 'http://trapias.github.io/assets/TrelloExport/default.css';
    }
    var htmlBody = '<!DOCTYPE html>\r\n<html><head><link type="text/css" rel="stylesheet" href="' + css + '"></head><body class="TrelloExport">\r\n' + html + '\r\n</body></html>';

    var now = new Date();
    var fileName = "TrelloExport_" + now.getFullYear() + dd(now.getMonth() + 1) + dd(now.getUTCDate()) + dd(now.getHours()) + dd(now.getMinutes()) + dd(now.getSeconds()) + ".html";

    saveAs(new Blob([s2ab(htmlBody)], {
        type: "text/html;charset=utf-8"
    }), fileName);

    // window.open(URL.createObjectURL(new Blob([s2ab(html)], {
    //     type: "text/html;charset=utf-8"
    // })));

    console.log('Done exporting ' + fileName);

    $.growl.notice({
        title: "TrelloExport",
        message: 'Done. Downloading HTML file...',
        fixed: true
    });
}

function createOPMLExport(jsonComputedCards) {

    var now = new Date();
    var sXML = '<?xml version="1.0" encoding="utf-8"?><opml version="1.0">';
    sXML += '<head><title>TrelloExport</title><dateCreated>' + now.toISOString() + '</dateCreated></head><body>';

    var prevBoard = null,
        prevList = null,
        description = null;

    // loop jsonComputedCards
    jsonComputedCards.forEach(function(card) {

        if (prevBoard !== card.boardName) {
            if (prevBoard !== null) {
                sXML += '</outline>';
            }
            // add board
            sXML += '<outline text="' + escape4XML(card.boardName.trim()) + '">';
            prevBoard = card.boardName;
        }

        if (prevList !== card.listName) {
            if (prevList !== null) {
                sXML += '</outline>';
            }
            // add list
            sXML += '<outline text="' + escape4XML(card.listName.trim()) + '">';
            prevList = card.listName;
        }

        // std is description, omnioutliner uses _note
        sXML += '<outline text="' + escape4XML(card.title.trim()) + '" _note="' + escape4XML(card.cardDescription) + '" created="' + card.datetimeCreated + '" createdBy="' + card.memberCreator + '"' +
            (card.due !== '' ? ' due="' + new Date(card.due).toLocaleDateString() + ' ' + new Date(card.due).toLocaleTimeString() + '"' : '') +
            (card.datetimeDone !== '' ? ' completed="' + card.datetimeDone + '"' : '') +
            (card.memberDone !== '' ? ' completedBy="' + card.memberDone + '"' : '') +
            (card.completionTimeText !== '' ? ' elapse="' + card.completionTimeText + '"' : '') +
            '>';

        // checklists
        var i;
        if (card.jsonCheckLists.length > 0) {
            sXML += '<outline text="Checklists">';
            for (i = 0; i < card.jsonCheckLists.length; i++) {
                sXML += '<outline text="' + card.jsonCheckLists[i].name + '">';
                for (var ii = 0; ii < card.jsonCheckLists[i].items.length; ii++) {
                    if (card.jsonCheckLists[i].items[ii].completed === true) {
                        sXML += '<outline text="' + '[x] ' + card.jsonCheckLists[i].items[ii].name + ' ' + card.jsonCheckLists[i].items[ii].completedDate + '" completedBy="' + card.jsonCheckLists[i].items[ii].completedBy + '"></outline>';
                    } else {
                        sXML += '<outline text="' + '[ ] ' + card.jsonCheckLists[i].items[ii].name + '"></outline>';
                    }
                }
                sXML += '</outline>';
            }
            sXML += '</outline>';
        }

        // comments
        if (card.jsonComments.length > 0) {
            sXML += '<outline text="Comments">';
            for (i = 0; i < card.jsonComments.length; i++) {
                var d = card.jsonComments[i].date;
                sXML += '<outline text="' + escape4XML(card.jsonComments[i].text) + '" created="' + d.toLocaleDateString() + ' ' + d.toLocaleTimeString() + '" createdBy="' + escape4XML(card.jsonComments[i].memberCreator.fullName) + '"></outline>';
            }
            sXML += '</outline>';
        }

        // attachments
        if (card.jsonAttachments.length > 0) {
            sXML += '<outline text="Attachments">';
            for (i = 0; i < card.jsonAttachments.length; i++) {
                sXML += '<outline text="' + escape4XML(card.jsonAttachments[i].name) + '" size="' + Number(card.jsonAttachments[i].bytes / 1024).toFixed(2) + ' kb" url="' + escape4XML(card.jsonAttachments[i].url) + '"></outline>';
            }
            sXML += '</outline>';
        }

        sXML += '</outline>';
    });

    sXML += '</outline></outline>';
    sXML += '</body></opml>';

    var oParser = new DOMParser();
    var oDOM = oParser.parseFromString(sXML, "text/xml");

    var s = new XMLSerializer();
    var opml = s.serializeToString(oDOM);
    // console.log('OPML: ' + opml);

    var fileName = "TrelloExport_" + now.getFullYear() + dd(now.getMonth() + 1) + dd(now.getUTCDate()) + dd(now.getHours()) + dd(now.getMinutes()) + dd(now.getSeconds()) + ".opml";

    saveAs(new Blob([opml], {
        type: "text/xml;charset=utf-8"
    }), fileName);

    // window.open(URL.createObjectURL(new Blob([s2ab(html)], {
    //     type: "text/html;charset=utf-8"
    // })));

    console.log('Done exporting ' + fileName);

    $.growl.notice({
        title: "TrelloExport",
        message: 'Done. Downloading OPML file...',
        fixed: true
    });
}