/*!
 * TrelloExport
 * https://github.com/llad/trelloExport
 *
 * Credit:
 * Started from: https://github.com/Q42/TrelloScrum
 * 
 * Forked by @trapias (Alberto Velo)
 * https://github.com/trapias/trelloExport
 * Whatsnew:
	- export Trello Plus Spent and Estimate data
	- export checklists
	- export comments (see commentLimit, default 100)
	- export attachments
	- export votes
	- use updated version of xlsx.js, modified by me (escapeXML)
	- use updated version of jquery, 2.1.0
 */

// Variables
var $excel_btn,
    columnHeadings = ['List', 'Title', 'Link', 'Description', 'Checklists', 'Comments', 'Attachments', 'Votes', 'Spent', 'Estimate', 'Due', 'Members', 'Labels'],
	commentLimit = 100; // limit the number of comments to put in the spreadsheet

window.URL = window.webkitURL || window.URL;

// on DOM load
$(function () {
 //  $(window).load(function () { 
   
   //todo: fix loading when side menu is collapsed
   
   setTimeout(function(){
    // Look for clicks on the .js-share class, which is
    // the "Share, Print, Export..." link on the board header option list
    $('.js-share').on('mouseup', function () {
        setTimeout(addExportLink);
    });
   }, 1000);
   
   //bind when opening side menu
   $('a.sidebar-show-btn').on('click', function(){
   //console.log('clicked sidebar');
	$('.js-share').on('mouseup', function () {
        setTimeout(addExportLink);
    });
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

    // RegEx to find the Trello Plus Spent and Estimate (spent/estimate) in card titles
    var SpentEstRegex = /\(([0-9]+(\.[0-9]+)?)\/?([0-9]+(\.[0-9]+)?)\)?/;
	
	console.log('Start export...');
	
	//var dlg = "<div id=\"dialog\" title=\"Basic dialog\"><p>This is the default dialog which is useful for displaying information. The dialog window can be moved, resized and closed with the 'x' icon.</p></div>";
	//$(dlg).dialog();
	
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
			if(data.name.length>30)
				w.name = data.name.substr(0,30);
			else
				w.name = data.name;
            w.data = [];
            w.data.push([]);
            w.data[0] = columnHeadings;
            
            
            // Setup the archive list and cart worksheet            
            wArchived = file.worksheets[1]; 
			wArchived.name = 'Archived cards';
			/* if(data.name.length>11)
				wArchived.name = 'Archived ' + data.name.substr(0,11);
			else
				wArchived.name = 'Archived ' + data.name; */
			
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
					
				console.log('TITLE ' + title);
				var spent=0; var estimate=0;
				var checkListsText='';
				var commentsText=''; var commentCounter=0;
				var attachmentsText='';
				
				//Trello Plus Spent/Estimate
				//TODO: fix!!!
				var spentData = title.match(SpentEstRegex);
				if(spentData!=null)
				{
					/* console.log('spentData=' + spentData);
					console.log('spentData[1]=' + spentData[1]);
					console.log('spentData[2]=' + spentData[2]);
					console.log('spentData[3]=' + spentData[3]);
					console.log('spentData[4]=' + spentData[3]); */
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
					console.log('parse ' + checklists.length + ' checklists for this card');
					//var checkListsText='';
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

						console.log('parse ' + data.actions.length + ' comments for this card');
						$.each(data.actions, function(j, action){  
						
							 if(action.type == "commentCard" && commentCounter <= commentLimit){
								if(card.id == action.data.card.id){
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
					
					if(card.attachments)
					{
						console.log('parse ' + card.attachments.length + ' attachments for this card');
						$.each(card.attachments, function(j, attach){  
							 attachmentsText += '[' + attach.name + '] (' + attach.bytes + ') ' + attach.url + '\n';
						});
					}
					 
                    // Need to set dates to the Date type so xlsx.js sets the right datatype
                    if (due !== '' ){
                        var d = new Date(due);
                        due = d;
                    }
                    
                    var rowData = [
                            listName,
                            title,
							'https://trello.com/c/' + card.id,
                            card.desc,
							checkListsText,
							commentsText,
							attachmentsText,
							card.idMembersVoted.length,
                            spent,
							estimate,
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
        
		console.log('Prepare xlsx...');
		
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

		console.log('Done exporting ' + board_title + '.xlsx');

    });

}
