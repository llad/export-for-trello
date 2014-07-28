/*
 * TrelloExport
 * https://github.com/llad/trelloExport
 *
 * Credit:
 * Started from: https://github.com/Q42/TrelloScrum
 * 
 * Forked by @trapias (Alberto Velo)
 * https://github.com/trapias/trelloExport
 *
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
    idBoard=0;
	
// Variables
var $excel_btn,
    columnHeadings = ['List', 'Card #', 'Title', 'Link', 'Description', 'Checklists', 'Comments', 'Attachments', 'Votes', 'Spent', 'Estimate', 'Created', 'CreatedBy', 'Due', 'Done', 'DoneBy', 'Members', 'Labels'],
	commentLimit = 100, // limit the number of comments to put in the spreadsheet
    loaded=false;
	
window.URL = window.webkitURL || window.URL;

//console.log('TrelloExport 1.9.3');
   
function sheet_from_array_of_arrays(data, opts) {
	var ws = {};
	var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
	for(var R = 0; R != data.length; ++R) {
		for(var C = 0; C != data[R].length; ++C) {
			if(range.s.r > R) range.s.r = R;
			if(range.s.c > C) range.s.c = C;
			if(range.e.r < R) range.e.r = R;
			if(range.e.c < C) range.e.c = C;
			var cell = {v: data[R][C] };
			if(cell.v == null) continue;
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
  // see below for better implementation!
  String.prototype.startsWith = function (str){
  //console.log(this + (this.indexOf(str) == 0 ? ' STARTSWITH' : '' + ' DOESNOTSTARTWITH') + ' ' + str);
    return this.indexOf(str) == 0;
  };
}

function getCreateCardAction(idCard) {
  if (!actionsCreateCard){
  //console.log('getCreateCardAction ' + 'https://trello.com/1/boards/' + idBoard + '/actions?filter=createCard&limit=1000');
    $.ajax({
      url:'https://trello.com/1/boards/' + idBoard + '/actions?filter=createCard&limit=1000', 
      dataType:'json',
      async: false,
      success: function(actionsData) {
        actionsCreateCard = actionsData;
      }
    });
  }
  var query = Enumerable.From(actionsCreateCard)
    .Where(function(x){if(x.data.card){return x.data.card.id == idCard}})
    .ToArray();
  return query.length > 0 ? query[0] : false;
}

function getMoveCardAction(idCard, nameList) {
  if (!actionsMoveCard){
  //console.log('getMoveCardAction ' + 'https://trello.com/1/boards/' + idBoard + '/actions?filter=updateCard:idList&limit=1000');
    $.ajax({
      url:'https://trello.com/1/boards/' + idBoard + '/actions?filter=updateCard:idList&limit=1000', 
      dataType:'json',
      async: false,
      success: function(actionsData) {
        actionsMoveCard = actionsData;
      }
    });
  }
  var query = Enumerable.From(actionsMoveCard)
    .Where(function(x){if(x.data.card && x.data.listAfter){return x.data.card.id == idCard && x.data.listAfter.name == nameList}})
    .OrderByDescending(function(x){return x.date;})
    .ToArray();
  return query.length > 0 ? query[0] : false;
}

function createExcelExport() {

    // RegEx to find the Trello Plus Spent and Estimate (spent/estimate) in card titles
    var SpentEstRegex = /\(([0-9]+(\.[0-9]+)?)\/?([0-9]+(\.[0-9]+)?)\)?/;
	
	console.log('Start export...');
	
    $.getJSON($('a.js-export-json').attr('href'), function (data) {
		    idBoard = data.id;
		
            // Setup the active list and cart worksheet
            w=new Object();
			if(data.name.length>30)
				w.name = data.name.substr(0,30);
			else
				w.name = data.name;
            w.data = [];
            w.data.push([]);
            w.data[0] = columnHeadings;
            
            // Setup the archive list and cart worksheet            
		    wArchived=new Object();
			wArchived.name = 'Archived cards';
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
					
				console.log('Card #' + card.idShort + ' ' + title);
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
				if(spentData!=null)
				{
					spent = spentData[1];
					estimate = spentData[3];
					
					if(spent==undefined) spent=0;
					if(estimate==undefined) estimate=0;
					
					//console.log('SPENT ' + spent + ' / estimate ' + estimate);
				}
				else
				{
					//no spent info found, do nothing
					//console.log('UNSPENT ' + spent + ' / estimate ' + estimate);
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
                    var memberIDs = card.idMembers;
                    var memberInitials = [];
                    $.each(memberIDs, function (i, memberID){
                        $.each(data.members, function (key, member) {
                            if (member.id == memberID) {
                                //memberInitials.push(member.initials);
								memberInitials.push(member.username);
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
					
					//all checklists
					var checklists = [];
					$.each(card.idChecklists, function (i, idchecklist){
					    if (idchecklist) {
                            checklists.push(idchecklist);
                        }
                    });
                    				
					//parse checklists
//					console.log('parse ' + checklists.length + ' checklists for this card');
					$.each(checklists, function(i, checklistid){
					//console.log('PARSE ' + checklistid);
						 $.each(data.checklists, function (key, list) {
							var list_id = list.id;
							//console.log('found ' + list_id);
							if(list_id == checklistid)
							{
								//console.log('CHECKLIST ' + list.name);
								checkListsText += list.name + ':\n';
								//checkitems
								$.each(list.checkItems, function (i, item){
									if (item) {
										//console.log('item ' + item.name + ' in state ' + item.state);
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
						//Date.js parsing, in progress
						//var shortDate = Date.CultureInfo.formatPatterns.shortDate;

//						console.log('parse ' + data.actions.length + ' actions for this card');
						$.each(data.actions, function(j, action){  
						
							 if(action.type == "commentCard" && commentCounter <= commentLimit){
								if(card.id == action.data.card.id){
								commentCounter ++;
								//2013-08-08T06:57:18 
								//var sActionDate = action.date.substr(0,19); //.toString( shortDate );
								var sActionDate = action.date.substr(0,10) + ' ' + action.date.substr(11,8);
								
								//console.log('sActionDate ' + sActionDate);
								//var sdate = Date.parse(sActionDate);
//								console.log('action.data.text=' + action.data.text);
									
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
					
					if(card.attachments)
					{
//						console.log('parse ' + card.attachments.length + ' attachments for this card');
						$.each(card.attachments, function(j, attach){  
							 attachmentsText += '[' + attach.name + '] (' + attach.bytes + ') ' + attach.url + '\n';
						});
					}
					
					//pulled from https://github.com/bmccormack/export-for-trello/blob/5b2b8b102b98ed2c49241105cb9e00e44d4e1e86/trelloexport.js
					//Get member created and DateTime created
                    var query = Enumerable.From(data.actions)
                      .Where(function(x){if(x.data.card){return x.data.card.id == card.id && x.type=="createCard"}})
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
					
					 //Find out when the card was most recently moved to any list whose name starts with "Done"
                    var nameListDone = "Done";
                    var query = Enumerable.From(data.actions)
                      .Where(function(x){if (x.data.card && x.data.listAfter){var listAfterName = x.data.listAfter.name; return x.data.card.id == card.id && listAfterName.startsWith(nameListDone);}})
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
                            card.desc,
							checkListsText,
							commentsText,
							attachmentsText,
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
        
		console.log('Prepare xlsx...');
		var board_title = data.name;
		var wb = new Workbook(), ws = sheet_from_array_of_arrays(w.data), wsArchived =  sheet_from_array_of_arrays(wArchived.data);
		/* add worksheet to workbook */
		wb.SheetNames.push(board_title);
		wb.Sheets[board_title] = ws;
		wb.SheetNames.push("Archived");
		wb.Sheets["Archived"] = wsArchived;
		
		var wbout = XLSX.write(wb, {bookType:'xlsx', bookSST:true, type: 'binary'});
		saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), board_title + ".xlsx")
		
		$("a.close-btn")[0].click();

		console.log('Done exporting ' + board_title + '.xlsx');

    });

}
