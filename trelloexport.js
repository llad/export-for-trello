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
 */
 var $,
    byteString,
    xlsx,
    ArrayBuffer,
    Uint8Array,
    Blob,
    saveAs,
    actionsCreateCard,
    actionsMoveCard,
    idBoard = 0,
    nProcessedLists = 0,
    nProcessedCards = 0,
    $excel_btn,
    columnHeadings = ['List', 'Card #', 'Title', 'Link', 'Description', 'Checklists', 'Comments', 'Attachments', 'Votes', 'Spent', 'Estimate', 'Created', 'CreatedBy', 'Due', 'Done', 'DoneBy', 'Members', 'Labels'],
	dataLimit = 1000, // limit the number of items retrieved from Trello
    loaded = false,
	MAXCHARSPERCELL=32767; /* excel limit */
	
// window.URL = window.webkitURL || window.URL;

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

function getCreateCardAction(idCard) {
  if (!actionsCreateCard){
  // console.log('getCreateCardAction ' + 'https://trello.com/1/boards/' + idBoard + '/actions?filter=createCard&limit=1000');
  // todo: make limit configurable
    $.ajax({
      url:'https://trello.com/1/boards/' + idBoard + '/actions?filter=createCard&limit=' + dataLimit, 
      dataType:'json',
      async: false,
      success: function(actionsData) {
        actionsCreateCard = actionsData;
      }
    });
  }
  var query = Enumerable.From(actionsCreateCard)
    .Where(function(x){if(x.data.card){return x.data.card.id == idCard;}})
    .ToArray();
  return query.length > 0 ? query[0] : false;
}

function getMoveCardAction(idCard, nameList) {
  if (!actionsMoveCard){
  //console.log('getMoveCardAction ' + 'https://trello.com/1/boards/' + idBoard + '/actions?filter=updateCard:idList&limit=1000');
    $.ajax({
      url:'https://trello.com/1/boards/' + idBoard + '/actions?filter=updateCard:idList&limit=' + dataLimit, 
      dataType:'json',
      async: false,
      success: function(actionsData) {
        actionsMoveCard = actionsData;
      }
    });
  }
  var query = Enumerable.From(actionsMoveCard)
    .Where(function(x){if(x.data.card && x.data.listAfter){return x.data.card.id == idCard && x.data.listAfter.name == nameList;}})
    .OrderByDescending(function(x){return x.date;})
    .ToArray();
  return query.length > 0 ? query[0] : false;
}

function TrelloExportOptions() {
	// console.log('TrelloExportOptions');

	// todo: open config / options dialog, then launch
	
	return createExcelExport();
	
}

function createExcelExport() {
	console.log('TrelloExport exporting to Excel...');
	  $.growl({  title: "TrelloExport", message: "Starting export" });

    // RegEx to find the Trello Plus Spent and Estimate (spent/estimate) in card titles
    var SpentEstRegex = /\(([0-9]+(\.[0-9]+)?)\/?([0-9]+(\.[0-9]+)?)\)?/;
	
	/*
		get data via Trello API instead of Board's json 
	*/
	var boardExportURL = $('a.js-export-json').attr('href');
	var parts = /\/b\/(\w{8})\.json/.exec(boardExportURL); // extract board id
	if(!parts) {
		$.growl.error({  title: "TrelloExport", message: "Board menu not open?" });
		return;
	}
	idBoard = parts[1];
	var apiURL = "https://trello.com/1/boards/" + idBoard + "?lists=all&cards=all&card_fields=all&card_checklists=all&members=all&member_fields=all&membersInvited=all&checklists=all&organization=true&organization_fields=all&fields=all&actions=commentCard%2CcopyCommentCard&card_attachments=true";

	$.getJSON(apiURL, function (data) {
		$.growl({  title: "TrelloExport", message: "Got data from Trello, parsing..." });
        // Setup the active list and cart worksheet
        w={}; //new Object();
		if(data.name.length>30)
			w.name = data.name.substr(0,30);
		else
			w.name = data.name;
        w.data = [];
        w.data.push([]);
        w.data[0] = columnHeadings;
        
        // Setup the archive list and cart worksheet            
	    wArchived={}; //new Object();
		wArchived.name = 'Archived cards';
		wArchived.data = [];
        wArchived.data.push([]);
        wArchived.data[0] = columnHeadings;
        
        // This iterates through each list and builds the dataset                     
        $.each(data.lists, function (key, list) {
            var list_id = list.id;
            var listName = list.name;                                            
            nProcessedLists++;
            // tag archived lists
            if (list.closed) {
                listName = '[archived] ' + listName;
            }
            
            // Iterate through each card and transform data as needed
            $.each(data.cards, function (i, card) {
            if (card.idList == list_id) {
                var title = card.name;
				
			console.log('Card #' + card.idShort + ' ' + title);
			nProcessedCards++;

			var spent=0; var estimate=0;
			var checkListsText='',
				commentsText='', 
				commentCounter=0,
				attachmentsText='',
				memberCreator='',
				datetimeCreated=null,
				memberDone,
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
									// console.log('item ' + item.name + ' in state ' + item.state + ', POS ' + item.pos);
									checkListsText += ' - ' + item.name + ' (' + item.state + ')\n';
								}
							});
							//checkListsText += '\n';
						}						
					});
				});

				//comments
				if(data.actions)
				{
					// console.log('Actions: ' + data.actions.length);
					//Date.js parsing, in progress
					//var shortDate = Date.CultureInfo.formatPatterns.shortDate;

					// console.log('parse ' + data.actions.length + ' actions for this card');
					$.each(data.actions, function(j, action){  
					
						 if((action.type == "commentCard" || action.type == 'copyCommentCard' )){
							if(card.id == action.data.card.id) {
								commentCounter ++;
								//2013-08-08T06:57:18 
								//var sActionDate = action.date.substr(0,19); //.toString( shortDate );
								var sActionDate = action.date.substr(0,10) + ' ' + action.date.substr(11,8);
								//console.log('sActionDate ' + sActionDate);
								//var sdate = Date.parse(sActionDate);

								if(action.memberCreator)
								{
									commentsText += '[' + sActionDate + ' - ' + action.memberCreator.username + '] ' + action.data.text + "\n";
								}
								else{
									commentsText += '[' + sActionDate + '] ' + action.data.text + "\n";
								}
							}
						} 
					});
				}

				// console.log('Found ' + commentCounter + ' comments');
				
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
                  datetimeCreated = query[0].date;
                }
                else {
                  //use the API to get the action created method
                  var actionCreateCard = getCreateCardAction(card.id);
                  if (actionCreateCard){
                    memberCreator = actionCreateCard.memberCreator.fullName;
                    datetimeCreated = actionCreateCard.date;
                  }
                  else {
                    memberCreator = "";
                    datetimeCreated = "";
                  }
                }
				
				 //Find out when the card was most recently moved to any list whose name starts with "Done" (ignore case, e.g. 'done' or 'DONE' or 'DoNe')
                var nameListDone = "Done";
                query = Enumerable.From(data.actions)
                  .Where(function(x){if (x.data.card && x.data.listAfter){var listAfterName = x.data.listAfter.name; return x.data.card.id == card.id && listAfterName.toLowerCase().startsWith(nameListDone.toLowerCase());}})
                  .OrderByDescending(function(x){return x.date;})
                  .ToArray();
                if (query.length > 0){
                  memberDone = query[0].memberCreator.fullName;
                  datetimeDone = query[0].date;
                }
                else {
                  var actionMoveCard = getMoveCardAction(card.id, nameListDone);
                  if (actionMoveCard){
                    memberDone = actionMoveCard.memberCreator.fullName;
                    datetimeDone = actionMoveCard.date;
                  }
                  else {
                    memberDone = "";
                    datetimeDone = "";
                  }
                }
				
                var rowData = [
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
        
        $.growl({  title: "TrelloExport", message: "Preparing xlsx..." });
		console.log('Prepare xlsx...');
		var board_title = data.name;
		var wb = new Workbook(), ws = sheet_from_array_of_arrays(w.data), wsArchived =  sheet_from_array_of_arrays(wArchived.data);
		/* add worksheet to workbook */
		wb.SheetNames.push(board_title);
		wb.Sheets[board_title] = ws;
		
		if(wArchived!==undefined)
		{
			wsArchived =  sheet_from_array_of_arrays(wArchived.data);
			 wb.SheetNames.push("Archived");
			wb.Sheets.Archived = wsArchived;
		}
		
		var wbout = XLSX.write(wb, {bookType:'xlsx', bookSST:true, type: 'binary'});
		saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), board_title + ".xlsx");
		
		$("a.pop-over-header-close-btn")[0].click();
		console.log('Processed ' + nProcessedLists + ' lists and ' + nProcessedCards + ' cards');
		console.log('Done exporting ' + board_title + '.xlsx');
		$.growl.notice({  title: "TrelloExport", message: 'Done. Processed ' + nProcessedLists + ' lists and ' + nProcessedCards + ' cards' , static: true});

    });

}
