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
 */
 var $,
    byteString,
    xlsx,
    ArrayBuffer,
    Uint8Array,
    Blob,
    saveAs,
    actionsCreateCard = [],
    actionsMoveCard = [],
    actionsCommentCard = [],
    idBoard = 0,
    nProcessedBoards = 0,
    nProcessedLists = 0,
    nProcessedCards = 0,
    $excel_btn,
    columnHeadings = ['Board', 'List', 'Card #', 'Title', 'Link', 'Description', 'Checklists', 'Comments', 'Attachments', 'Votes', 'Spent', 'Estimate', 'Created', 'CreatedBy', 'Due', 'Done', 'DoneBy', 'DoneTime', 'Members', 'Labels'],
    dataLimit = 1000, // limit the number of items retrieved from Trello
    MAXCHARSPERCELL=32767,
    exportlists=[],
    exportboards=[],
    exportcards=[],
    nameListDone = "Done",
    filterListsNames = [];

function sheet_from_array_of_arrays(data, opts) {
    // console.log('sheet_from_array_of_arrays ' + data);
    var ws = {};
    var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
    for(var R = 0; R != data.length; ++R) {
        for(var C = 0; C != data[R].length; ++C) {
            if(range.s.r > R) range.s.r = R;
            if(range.s.c > C) range.s.c = C;
            if(range.e.r < R) range.e.r = R;
            if(range.e.c < C) range.e.c = C;
            var cell = {v: data[R][C] };
            if(cell.v === null) continue;
            var cell_ref = XLSX.utils.encode_cell({c:C,r:R});
            
            if(typeof cell.v === 'number') cell.t = 'n';
            else if(typeof cell.v === 'boolean') cell.t = 'b';
            else if(cell.v instanceof Date) {
                cell.t = 'n'; cell.z = XLSX.SSF._table[14];
                cell.v = datenum(cell.v);
            }
            else cell.t = 's';
            
            ws[cell_ref] = cell;
        }
    }
    if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
    return ws;
}

function Workbook() {
    if(!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}

function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}

if (typeof String.prototype.startsWith != 'function') {
  String.prototype.startsWith = function (str){
    return this.indexOf(str) === 0;
  };
}

function getCommentCardActions(boardID, idCard) {
    for(var n=0; n<actionsCommentCard.length;n++) {
        if(actionsCommentCard[n].board===boardID) {
            var query = Enumerable.From(actionsCommentCard[n].data)
            .Where(function(x){if(x.data.card){return x.data.card.id == idCard;}})
            .OrderByDescending(function(x){return x.date;})
            .ToArray();
          return query.length > 0 ? query : false;
        }
    }
    $.ajax({
      url:'https://trello.com/1/boards/' + boardID + '/actions?filter=commentCard,copyCommentCard&limit=' + dataLimit, 
      dataType:'json',
      async: false,
      success: function(actionsData) {
          var a = {
              board: boardID,
              data:  actionsData,
          };
        actionsCommentCard.push(a);
      }
    });
      var selectedActions = null;
      // get the right actions for board
    for(var n=0; n<actionsCommentCard.length;n++) {
        if(actionsCommentCard[n].board===boardID) {
            selectedActions = actionsCommentCard[n].data;
            break;
        }
    }
    var query = Enumerable.From(selectedActions)
        .Where(function(x){if(x.data.card){return x.data.card.id == idCard;}})
        .OrderByDescending(function(x){return x.date;})
        .ToArray();
    return query.length > 0 ? query : false;
}

function getCreateCardAction(boardID, idCard) {
    for(var n=0; n<actionsCreateCard.length;n++) {
        if(actionsCreateCard[n].board===boardID) {
            var query = Enumerable.From(actionsCreateCard[n].data)
            .Where(function(x){if(x.data.card){return x.data.card.id == idCard;}})
            .ToArray();
          return query.length > 0 ? query[0] : false;
        }
    }
    $.ajax({
      url:'https://trello.com/1/boards/' + boardID + '/actions?filter=createCard&limit=' + dataLimit, 
      dataType:'json',
      async: false,
      success: function(actionsData) {
        var a = {
              board: boardID,
              data:  actionsData,
          };
        actionsCreateCard.push(a);
      }
    });
    var selectedActions = null;
      // get the right actions for board
    for(var n=0; n<actionsCreateCard.length;n++) {
        if(actionsCreateCard[n].board===boardID) {
            selectedActions = actionsCreateCard[n].data;
            break;
        }
    }
  var query = Enumerable.From(selectedActions)
    .Where(function(x){if(x.data.card){return x.data.card.id == idCard;}})
    .ToArray();
  return query.length > 0 ? query[0] : false;
}

function getMoveCardAction(boardID, idCard, nameList) {
    for(var n=0; n<actionsMoveCard.length;n++) {
        if(actionsMoveCard[n].board===boardID) {
            var query = Enumerable.From(actionsMoveCard[n].data)
            .Where(function(x){if(x.data.card && x.data.listAfter){return x.data.card.id == idCard && x.data.listAfter.name == nameList;}})
            .OrderByDescending(function(x){return x.date;})
            .ToArray();
          return query.length > 0 ? query[0] : false;
        }
    }
    $.ajax({
      url:'https://trello.com/1/boards/' + boardID + '/actions?filter=updateCard&limit=' + dataLimit, 
      dataType:'json',
      async: false,
      success: function(actionsData) {
          var a = {
              board: boardID,
              data:  actionsData,
          };
        actionsMoveCard.push(a);
      }
    });
      var selectedActions = null;
      // get the right actions for board
    for(var n=0; n<actionsMoveCard.length;n++) {
        if(actionsMoveCard[n].board===boardID) {
            selectedActions = actionsMoveCard[n].data;
            break;
        }
    }
    var query = Enumerable.From(selectedActions)
        .Where(function(x){if(x.data.card && x.data.listAfter){return x.data.card.id == idCard && x.data.listAfter.name == nameList;}})
        .OrderByDescending(function(x){return x.date;})
        .ToArray();
    return query.length > 0 ? query[0] : false;
}

function searchupdateCheckItemStateOnCardAction(checkitemid, actions) {
    var sOut = '';
    $.each(actions, function(j, action){  
        if(action.type == 'updateCheckItemStateOnCard') {
            if(action.data.checkItem.id == checkitemid) {
                var d = new Date(action.date);
                var sActionDate = d.toLocaleDateString() + ' ' + d.toLocaleTimeString(); // .substr(0,10) + ' ' + action.date.substr(11,8);
                sOut = '[completed ' + sActionDate + ' by ' + action.memberCreator.fullName + ']';
                console.log('checkitemid ' + checkitemid + '=' + sOut);
                return sOut;
            }
        }
    });
    return sOut;
}

function TrelloExportOptions() {

    exportboards=[];
    exportlists=[]; // reset
    filterListsNames = [];
    nProcessedBoards = 0;
    nProcessedLists = 0;
    nProcessedCards = 0;

    var sDialog = '<table id="optionslist">' +
        '<tr><td>Max number of items retrieved from Trello:</td><td><input type="text" size="4" name="setdatalimit" id="setdatalimit" value="1000" /></td></tr>' + 
        '<tr><td>Done lists name:</td><td><input type="text" size="4" name="setnameListDone" id="setnameListDone" value="' + nameListDone + '" /></td></tr>' +
        '<tr><td>Type of export:</td><td><select id="exporttype"><option value="board">Current Board</option><option value="list">Select Lists in current Board</option><option value="boards">Multiple Boards</option><option value="cards">Select cards in a list</option></select></td></tr>' +
        '</table>';

    setTimeout(function() {

        $('#exporttype').on('change', function() {
            var sexporttype = $('#exporttype').val();
            switch(sexporttype) {
                case 'list':
                    $('#choosenboards').parent().parent().remove();
                    $('#choosenCards').parent().parent().remove();
                    $('#choosenSinglelist').parent().parent().remove();
                    // get a list of all lists in board and let user choose which to export
                    var sSelect = getalllistsinboard();
                    $('#optionslist').append('<tr><td>Select one or more Lists</td><td><select multiple id="choosenlist">' + sSelect + '</select></td></tr>');
                    break;
                case 'board':
                    $('#choosenlist').parent().parent().remove();
                    $('#choosenboards').parent().parent().remove();
                    $('#choosenCards').parent().parent().remove();
                    $('#choosenSinglelist').parent().parent().remove();
                    break;
                case 'boards':
                    $('#choosenlist').parent().parent().remove();
                    $('#choosenCards').parent().parent().remove();
                    $('#choosenSinglelist').parent().parent().remove();
                    // get a list of all boards
                    var sSelect = getallboards();
                    $('#optionslist').append('<tr><td>Select one or more Boards</td><td><select multiple id="choosenboards">' + sSelect + '</select></td></tr>');
                    $('#optionslist').append('<tr><td>Filter lists by name:</td><td><input type="text" size="4" name="filterListsNames" id="filterListsNames" value="" /></td></tr>');
                    break;
                case 'cards':
                    $('#choosenlist').parent().parent().remove();
                    $('#choosenboards').parent().parent().remove();
                    // get a list of all lists in board and let user choose which to export
                    var sSelect = getalllistsinboard();
                    $('#optionslist').append('<tr><td>Select one List</td><td><select id="choosenSinglelist"><option value="">Select a list</option>' + sSelect + '</select></td></tr>');

                    $('#choosenSinglelist').on('change', function() {
                        $('#choosenCards').parent().parent().remove();
                        var selectedList = $('#choosenSinglelist').val();
                        exportlists=[];
                        exportlists.push(selectedList);
                        // get card in list
                        var sSelect = getallcardsinlist(selectedList);
                        $('#optionslist').append('<tr><td>Select one or more cards</td><td><select multiple id="choosenCards">' + sSelect + '</select></td></tr>');
                    });
                    break;
                default:
                    break;
            }

        });

    }, 500);

    // open options dialog to configure & launch export
    $.Zebra_Dialog(sDialog, {
        title: 'TrelloExport Options',
        type: false,
        'buttons':  [
                    {
                        caption: 'Export', callback: function() { 

                            // dataLimit
                            var sLimit = $('#setdatalimit').val();
                            if($.isNumeric(sLimit)) {
                                dataLimit = sLimit;
                                if(dataLimit<1) {
                                    alert('Invalid datalimit, please specify a positive limit');
                                    $('#setdatalimit').val('1000');
                                    return false;    
                                }
                            } else {
                                alert('Invalid datalimit, please specify a numeric value');
                                $('#setdatalimit').val('1000');
                                return false;
                            }

                            nameListDone = $('#setnameListDone').val();

                            // export type
                            var sexporttype = $('#exporttype').val();
                            switch(sexporttype) {
                                case 'list':
                                    if( $('#choosenlist').length <= 0) {
                                        // console.log('wait for lists to load');
                                        return false;
                                    } else {
                                        $('#choosenlist > option:selected').each(function() {
                                            exportlists.push($(this).val());
                                        });                                        
                                    }
                                break;

                                case 'boards':
                                    if( $('#choosenboards').length <= 0) {
                                        // console.log('wait for lists to load');
                                        return false;
                                    } else {
                                        var sfilterListsNames = $('#filterListsNames').val();
                                        if(sfilterListsNames.trim()!=='') {
                                            // parse list name filters
                                            var filters = sfilterListsNames.split(',');
                                            for(var nd=0; nd<filters.length;nd++){
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
                                    if( $('#choosenCards').length <= 0) {
                                        // console.log('wait for cards to load');
                                        return false;
                                    } else {
                                        $('#choosenCards > option:selected').each(function() {
                                            exportcards.push($(this).val());
                                        });                                        
                                    }
                                break;

                                default: 
                                break;
                            }

                            // console.log('datalimit=' + dataLimit + ', sexporttype=' + sexporttype + ', exportlists=' + exportlists);
                           
                           // launch export
                           return createExcelExport();
                    }},
                    {
                        caption: 'Cancel', callback: function() { return; } // close dialog
                    }]
    });
    
    return; // close dialog
}

function getalllistsinboard() {

    var boardExportURL = $('a.js-export-json').attr('href');
    var parts = /\/b\/(\w{8})\.json/.exec(boardExportURL); // extract board id
    if(!parts) {
        $.growl.error({  title: "TrelloExport", message: "Board menu not open?" });
        return;
    }
    idBoard = parts[1];
    var apiURL = "https://trello.com/1/boards/" + idBoard + "?lists=all&cards=none";
    var sHtml = "";

    $.ajax({
        url: apiURL,
        async: false,
    })
    .done(function(data) {
         // console.log('DATA:' + JSON.stringify(data));
        $.each(data.lists, function (key, list) {
            var list_id = list.id;
            var listName = list.name;
            if(!list.closed) {
                sHtml += '<option value="' + list_id + '">' + listName + '</option>';
            } else {
                sHtml += '<option value="' + list_id + '">' + listName + ' [Archived]</option>';
            }
        });
    })
    .fail(function() {
        console.log("error");
    })
    .always(function() {
        // console.log("complete");
    });
    
    return sHtml;
}

function getorganizationid() {
    var boardExportURL = $('a.js-export-json').attr('href');
    var parts = /\/b\/(\w{8})\.json/.exec(boardExportURL); // extract board id
    if(!parts) {
        $.growl.error({  title: "TrelloExport", message: "Board menu not open?" });
        return;
    }
    idBoard = parts[1];
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
    .fail(function() {
        console.log("error");
    })
    .always(function() {
        // console.log("complete");
    });
    
    return orgID;
}

function getallboards() {

    var orgID = getorganizationid();

    // GET /1/organizations/[idOrg or name]/boards
    var apiURL = "https://trello.com/1/organizations/" + orgID + "/boards?lists=none";
    var sHtml = "";

    $.ajax({
        url: apiURL,
        async: false,
    })
    .done(function(data) {
         for (var i = 0; i < data.length; i++) {
                 var board_id = data[i].id;
                 var boardName = data[i].name;
                 if(!data[i].closed) {
                    sHtml += '<option value="' + board_id + '">' + boardName + '</option>';
                } else {
                    sHtml += '<option value="' + board_id + '">' + boardName + ' [Archived]</option>';
                }
         }
    })
    .fail(function() {
        console.log("error");
    })
    .always(function() {
        // console.log("complete");
    });
    
    return sHtml;
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
                 if(!data[i].closed) {
                    sHtml += '<option value="' + card_id + '">' + cardName + '</option>';
                } else {
                    sHtml += '<option value="' + card_id + '">' + cardName + ' [Archived]</option>';
                }
         }
    })
    .fail(function() {
        console.log("error");
    })
    .always(function() {
        // console.log("complete");
    });
    
    return sHtml;
}

function createExcelExport() {
    console.log('TrelloExport exporting to Excel...');
      $.growl({  title: "TrelloExport", message: "Starting export" });

    // RegEx to find the Trello Plus Spent and Estimate (spent/estimate) in card titles
    var SpentEstRegex = /\(([0-9]+(\.[0-9]+)?)\/?([0-9]+(\.[0-9]+)?)\)?/;
    
    /*
        get data via Trello API instead of Board's json 
    */
    
    if(exportboards.length===0) {
        // export just current board
        var boardExportURL = $('a.js-export-json').attr('href');
        var parts = /\/b\/(\w{8})\.json/.exec(boardExportURL); // extract board id
        if(!parts) {
            $.growl.error({  title: "TrelloExport", message: "Board menu not open?" });
            return;
        }
        idBoard = parts[1];
        exportboards.push(idBoard);
    }

    var wb = new Workbook();
    wArchived={};
    wArchived.name = 'Archived lists and cards';
    wArchived.data = [];
    wArchived.data.push([]);
    wArchived.data[0] = columnHeadings;


    // loop boards
    for (var iBoard = 0; iBoard < exportboards.length; iBoard++) {
        console.log('Export board ' + exportboards[iBoard]);
        idBoard = exportboards[iBoard];
        var apiURL = "https://trello.com/1/boards/" + idBoard + "?lists=all&cards=all&card_fields=all&card_checklists=all&members=all&member_fields=all&membersInvited=all&checklists=all&organization=true&organization_fields=all&fields=all&actions=commentCard%2CcopyCommentCard%2CupdateCheckItemStateOnCard&card_attachments=true";
        $.ajax({
                url: apiURL,
                async: false,
            })
            .done(function(data) {
            
            $.growl({  title: "TrelloExport", message: "Got data from Trello, parsing board " + idBoard + "..." });

            // Setup the active list and cart worksheet
            w={}; //new Object();
            if(data.name.length>30)
                w.name = data.name.substr(0,30);
            else
                w.name = data.name;
            w.data = [];
            w.data.push([]);
            w.data[0] = columnHeadings;

            // This iterates through each list and builds the dataset                     
            $.each(data.lists, function (key, list) {
                var list_id = list.id;
                var listName = list.name;
                
                if(exportlists.length>0) {
                    
                    if ($.inArray(list_id, exportlists) === -1)
                    {
                        console.log('skip list ' + listName);
                        return true;
                    }
                }

                // 1.9.14: filter lists by name
                var accept = true;
                if(filterListsNames.length>0) {
                    for(var y=0; y < filterListsNames.length; y++) {
                        if(!listName.toLowerCase().startsWith(filterListsNames[y].trim().toLowerCase())) {
                            accept = false;
                        } else {
                            accept = true;
                            break;
                        }
                    }
                }
                if(!accept) {
                    console.log('skipping list ' + listName);
                    return true;
                }

                console.log('processing list ' + listName);
                nProcessedLists++;

                // tag archived lists
                if (list.closed) {
                    listName = '[archived] ' + listName;
                }
                
                // Iterate through each card and transform data as needed
                $.each(data.cards, function (i, card) {
                if (card.idList == list_id) {
                    
                    //export selected cards only option
                    if(exportcards.length>0) {
                        if ($.inArray(card.id, exportcards) === -1)
                        {
                            console.log('skip card ' + card.id);
                            return true;
                        }
                    }

                    var title = card.name;
                    
                console.log('Card #' + card.idShort + ' ' + title);
                nProcessedCards++;

                var spent=0; var estimate=0;
                var checkListsText='',
                    commentsText='', 
                    attachmentsText='',
                    memberCreator='',
                    datetimeCreated=null,
                    memberDone='',
                    datetimeDone ;
                
                //Trello Plus Spent/Estimate
                var spentData = title.match(SpentEstRegex);
                if(spentData!==null)
                {
                    spent = spentData[1];
                    estimate = spentData[3];
                    if(spent===undefined) spent=0;
                    if(estimate===undefined) estimate=0;
                    // console.log('SPENT ' + spent + ' / estimate ' + estimate);
                }
                else
                {
                    //no spent info found, do nothing
                    // console.log('UNSPENT ' + spent + ' / estimate ' + estimate);
                    spent=0;
                    estimate=0;
                }
                
                title = title.replace(SpentEstRegex, '');
                title = title.replace('() ', '');
                    
                    // tag archived cards
                    if (card.closed) {
                        title = '[archived] ' + title;
                    }
                    var due = card.due || '';
                    
                    //Get all the Member IDs
                    // console.log('Members: ' + card.idMembers.length);
                    var memberIDs = card.idMembers;
                    var memberInitials = [];
                    $.each(memberIDs, function (i, memberID){
                        $.each(data.members, function (key, member) {
                            if (member.id == memberID) {
                                memberInitials.push(member.fullName); // initials, username or fullName
                            }
                        });
                    });
                    
                    //Get all labels
                    // console.log('Labels: ' + card.labels.length);
                    var labels = [];
                    $.each(card.labels, function (i, label){
                        if (label.name) {
                            labels.push(label.name);
                        }
                        else {
                            labels.push(label.color);
                        }

                    });
                    
                    //all checklists
                    // console.log('Checklists: ' + card.idChecklists.length);
                    var checklists = [];
                    if(card.idChecklists!==undefined)
                        $.each(card.idChecklists, function (i, idchecklist){
                            if (idchecklist) {
                                checklists.push(idchecklist);
                            }
                        });
                                    
                    //parse checklists
                    $.each(checklists, function(i, checklistid){
                    // console.log('PARSE ' + checklistid);
                         $.each(data.checklists, function (key, list) {
                            var list_id = list.id;
                            //console.log('found ' + list_id);
                            if(list_id == checklistid)
                            {
                                checkListsText += list.name + ':\n';
                                //checkitems: reordered (issue #4 https://github.com/trapias/trelloExport/issues/4)
                                var orderedChecklists = Enumerable.From(list.checkItems)
                                  .OrderBy(function(x){return x.pos;})
                                  .ToArray();

                                $.each(orderedChecklists, function (i, item){
                                    if (item) {
                                        if(item.state=='complete') {
                                            // issue #5
                                            // find who and when item was completed
                                            var sCompletedBy = searchupdateCheckItemStateOnCardAction(item.id, data.actions);
                                            checkListsText += ' - ' + item.name + ' ' + sCompletedBy + '\n';
                                        } else {
                                            checkListsText += ' - ' + item.name + ' [' + item.state + ']\n';
                                        }
                                    }
                                });
                            }                        
                        });
                    });

                    //comments
                    var commentsOnCard = getCommentCardActions(idBoard, card.id);
                    if(commentsOnCard)
                    {
                        // console.log('parse ' + data.actions.length + ' actions for this card');
                        $.each(commentsOnCard, function(j, action) {
                             if((action.type == "commentCard" || action.type == 'copyCommentCard' )){
                                if(card.id == action.data.card.id) {
                                    //2013-08-08T06:57:18 
                                    var d = new Date(action.date);
                                    var sActionDate = d.toLocaleDateString() + ' ' + d.toLocaleTimeString(); //.substr(0,10) + ' ' + action.date.substr(11,8);
                                    if(action.memberCreator)
                                    {
                                        commentsText += '[' + sActionDate + ' - ' + action.memberCreator.fullName + '] ' + action.data.text + "\n";
                                    }
                                    else{
                                        commentsText += '[' + sActionDate + '] ' + action.data.text + "\n";
                                    }
                                }
                            }
                        });
                    }

                    if(card.attachments)
                    {
                        // console.log('Attachments: ' + card.attachments.length);
                        $.each(card.attachments, function(j, attach){  
                            // console.log(attach);
                             attachmentsText += '[' + attach.name + '] (' + attach.bytes + ') ' + attach.url + '\n';
                        });
                    }
                    
                    //pulled from https://github.com/bmccormack/export-for-trello/blob/5b2b8b102b98ed2c49241105cb9e00e44d4e1e86/trelloexport.js
                    //Get member created and DateTime created
                    var query = Enumerable.From(data.actions)
                      .Where(function(x){if(x.data.card){return x.data.card.id == card.id && x.type=="createCard";}})
                      .ToArray();
                    if (query.length > 0){
                      memberCreator = query[0].memberCreator.fullName + ' (' + query[0].memberCreator.username + ')';
                      datetimeCreated = new Date(query[0].date);
                    }
                    else {
                      //use the API to get the action created method
                      var actionCreateCard = getCreateCardAction(idBoard, card.id);
                      if (actionCreateCard){
                        memberCreator = actionCreateCard.memberCreator.fullName;
                        datetimeCreated = new Date(actionCreateCard.date);
                      }
                      else {
                          // calculate datetimeCreated from card id
                          // cfr http://help.trello.com/article/759-getting-the-time-a-card-or-board-was-created
                          datetimeCreated = new Date(1000*parseInt(card.id.substring(0,8),16));
                        memberCreator = "";
                      }
                    }
                    
                    /**
                     * 1.9.14: handle multiple nameListDone
                     * e.g. Done, Finished
                     */
                    if(nameListDone==='') {nameListDone='Done';} // default
                    var allnameListDone = nameListDone.split(',');
                    for(var nd=0; nd<allnameListDone.length;nd++){
                        // console.log(nd + ') ' + allnameListDone[nd]);
                        if(memberDone==="") {
                            //Find out when the card was most recently moved to any list whose name starts with "Done" (ignore case, e.g. 'done' or 'DONE' or 'DoNe')
                            query = Enumerable.From(data.actions)
                              .Where(function(x){
                                  if (x.data.card && x.data.listAfter) {
                                      var listAfterName = x.data.listAfter.name; 
                                      return x.data.card.id == card.id && listAfterName.toLowerCase().startsWith(allnameListDone[nd].trim().toLowerCase());}
                              })
                              .OrderByDescending(function(x){return x.date;})
                              .ToArray();
                            if (query.length > 0) {
                                memberDone = query[0].memberCreator.fullName;
                                  datetimeDone = query[0].date;
                            } else {
                                var actionMoveCard = getMoveCardAction(idBoard, card.id, allnameListDone[nd].trim());
                                if (actionMoveCard) {
                                    memberDone = actionMoveCard.memberCreator.fullName;
                                    datetimeDone = actionMoveCard.date;
                                } else {
                                    memberDone = "";
                                    datetimeDone = "";
                                }
                            }
                        }
                    }
                    
                    var completionTime = "";
                    if (datetimeDone!="" && datetimeCreated!="") {
                        // var d1 = new Date(datetimeCreated);
                        var d2 = new Date(datetimeDone);
                        var df = new DateDiff(d2,datetimeCreated);
                        // PnYnMnDTnHnMnS ISO8601 -> PnDTnHnMnS
                        completionTime = "P" + df.days + "DT" + df.hours + "H" + df.minutes + "M" + df.seconds + "S";
                        datetimeCreated = datetimeCreated.toLocaleDateString() + ' ' + datetimeCreated.toLocaleTimeString();
                        datetimeDone = d2.toLocaleDateString() + ' ' + d2.toLocaleTimeString();
                    } else {
                        if(datetimeCreated) {
                            datetimeCreated = datetimeCreated.toLocaleDateString() + ' ' + datetimeCreated.toLocaleTimeString();    
                        }
                    }

                    var rowData = [
                            data.name,
                            listName,
                            card.idShort,
                            title,
                            'https://trello.com/c/' + card.shortLink,
                            card.desc.substr(0,MAXCHARSPERCELL),
                            checkListsText.substr(0,MAXCHARSPERCELL),
                            commentsText.substr(0,MAXCHARSPERCELL),
                            attachmentsText.substr(0,MAXCHARSPERCELL),
                            card.idMembersVoted.length,
                            spent,
                            estimate,
                            datetimeCreated,
                            memberCreator,
                            due,
                            datetimeDone,
                            memberDone,
                            completionTime,
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

         // $.growl({  title: "TrelloExport", message: "Preparing xlsx..." });
        // console.log('Prepare xlsx...');
        var board_title = data.name;
        console.log('board_title = ' + board_title);
        // var wb = new Workbook(),
        var ws = sheet_from_array_of_arrays(w.data);
        /* add worksheet to workbook */
        wb.SheetNames.push(board_title);
        wb.Sheets[board_title] = ws;
        // console.log("Added sheet " + board_title);
        })
        .fail(function() {
            console.log("error");
        });

        // end loop boards
        nProcessedBoards++;
    }

    //add the Archived data
    var wsArchived =  sheet_from_array_of_arrays(wArchived.data);
    if(wsArchived!==undefined)
    {
        wb.SheetNames.push("Archive");
        wb.Sheets.Archive = wsArchived;
    }

    var now = new Date();
    var fileName = "TrelloExport_" + now.getFullYear() + now.getMonth() + now.getUTCDate() + now.getHours() + now.getMinutes() + now.getSeconds() + ".xlsx";
    var wbout = XLSX.write(wb, {bookType:'xlsx', bookSST:true, type: 'binary'});
    saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), fileName);
    
    $("a.pop-over-header-close-btn")[0].click();
    console.log('Processed ' + nProcessedLists + ' lists and ' + nProcessedCards + ' cards');
    console.log('Done exporting ' + fileName);
    $.growl.notice({  title: "TrelloExport", message: 'Done. Processed ' + nProcessedLists + ' lists and ' + nProcessedCards + ' cards in ' + nProcessedBoards + (nProcessedBoards > 1 ? ' boards.' : ' board.') , static: true});
     
    console.log("The End");

}
