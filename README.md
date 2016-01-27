TrelloExport
============

Chrome extension to export data from Trello to Excel. And more to come.

TrelloExport now has its own Trello board, go to [https://trello.com/b/MBnwUMwM/trelloexport](https://trello.com/b/MBnwUMwM/trelloexport).

## Fork
This is a fork of the original TrelloExport extension, available at [https://github.com/llad/export-for-trello](https://github.com/llad/export-for-trello).

### Initial Fork modifications
This fork initially added the following modifications:

1. export Trello Plus Spent and Estimate values, extracting values from card titles in format (S/E)
2. export comments (with a default limit of 100, see commentLimit in script)
3. export checklists and checklists' items
4. export list of attachments
5. export a link to each card
6. export votes
7. use updated version of xlsx.js, modified by me (escapeXML)
8. use updated version of jquery, 2.1.0
9. use usernames instead of initials for members

Other improvements have been added since then: see the version history below (newer versions up).

### Version 1.9.15
Finally fixing comments export: should have finally fixed exporting of comments and 'done time' calculations: thanks @fepsch for sharing a board and allowing to identify this annoying bug.

### Version 1.9.14
Some bugfix and some new features.

**Fixes**:

- fixed card completion calculation when exporting multiple boards (fix getMoveCardAction and getCreateCardAction)
- loading comments with a new function (getCommentCardActions), trying to fix issues with comments reported by some users; please give feedback

**New features**:

- formatting dates in user (browser) locale
- added support for multiple 'Done' list names
- added capability to optionally filter exported lists by name when exporting multiple boards

Both the 'Done list name' and 'Filter lists by name' input boxes accept a comma-separated list of partial list names, i.e. just specify multiple names in the textbox like 'Done,Completed' (without apices). Lists will then be (case insensitively) matched when their name starts with one of these values.

More info in this [blog post](http://trapias.github.io/blog/trelloexport-1-9-14/). 

### Version 1.9.13:
Some (interesting, hopefully!) improvements with this version:

- new 'DoneTime' column holding card completion time in days, hours, minutes and seconds, formatted as per [ISO8601](https://en.wikipedia.org/wiki/ISO_8601)
- name (prefix) of 'Done' lists is now configurable, default "Done"
- larger options dialog to better show options
- export multiple (selected) boards
- export multiple (selected) cards in a list (i.e. export single cards)

More info in this [blog post](http://trapias.github.io/blog/trelloexport-1-9-13/). Give feeback!


### Version 1.9.12:
Fixed a bug by which the previously used BoardID was kept when navigating to another board.


### Version 1.9.11:

- added a new Options dialog
- export full board or choosen list(s) only
- add who and when item was completed to checklist items as of [issue #5](https://github.com/trapias/trelloExport/issues/5)

More info in this [blog post](http://trapias.github.io/blog/trelloexport-1-9-11/).

Your feedback is welcome, just comment on the blog, on the dedicated [Trello board](https://trello.com/b/MBnwUMwM/trelloexport) or open new issues.


### Version 1.9.10:

- adapt inject script to modified Trello layout

### Version 1.9.9:

- MAXCHARSPERCELL limit to avoid import errors in Excel (see https://support.office.com/en-nz/article/Excel-specifications-and-limits-16c69c74-3d6a-4aaf-ba35-e6eb276e8eaa)
- removed commentLimit, all comments are loaded (but attention to MAXCHARSPERCELL limit above, since comments go to a single cell)
- growl notifications with jquery-growl http://ksylvest.github.io/jquery-growl/

### Version 1.9.8:
Use Trello API to get data, thanks https://github.com/mjearhart and https://github.com/llad:

- https://github.com/llad/export-for-trello/pull/20
- https://github.com/mjearhart/export-for-trello/commit/2a07561fdcdfd696dee0988cbe414cfd8374b572

### Version 1.9.7:
- fix issue #3 (copied comments missing in export)

### Version 1.9.6
- order checklist items by position (issue #4)
- minor code changes

### Version 1.9.5
- code lint
- ignore case in finding 'Done' lists (thanks [AlvonsiusAlbertNainupu](https://disqus.com/by/AlvonsiusAlbertNainupu/))

### Version 1.9.4
Fixed bug preventing export when there are no archived cards.

### Version 1.9.3
Whatsnew for version 1.9.3:

- restored archived cards sheet

### Version 1.9.2
Whatsnew for version 1.9.2:

- fixed blocking error when duedate specified - thanks @ggyaniv for help
- new button loading function: the "Export Excel" button should always appear now

### Version 1.9.1
Whatsnew for version 1.9.1:

- fixed button loading
- some code cleaning

### Version 1.9.0
Whatsnew for version 1.9.0

- switched to SheetJS library to export to excel, cfr [https://github.com/SheetJS/js-xlsx](https://github.com/SheetJS/js-xlsx "https://github.com/SheetJS/js-xlsx")
- unicode characters are now correctly exported to xlsx


### Version 1.8.9
Whatsnew for version 1.8.9:

- added column Card #
- added columns memberCreator, datetimeCreated, datetimeDone and memberDone pulling modifications from [https://github.com/bmccormack/export-for-trello/blob/5b2b8b102b98ed2c49241105cb9e00e44d4e1e86/trelloexport.js](https://github.com/bmccormack/export-for-trello/blob/5b2b8b102b98ed2c49241105cb9e00e44d4e1e86/trelloexport.js "https://github.com/bmccormack/export-for-trello/blob/5b2b8b102b98ed2c49241105cb9e00e44d4e1e86/trelloexport.js")
- added linq.min.js library to support linq queries for the above modifications

---

#### Notes
I modified the **escapeXML** function in **xlsx.js** to avoid errors with XML characters when loading the spreadsheet in Excel. I tested exporting quite big boards like Trello Development or Trello Resources and no more have issues with invalid characters.
I put a couple of sample export files in the xlsx subfolder.

**Columns**: the list of columns exported is now:

	columnHeadings = ['List', 'Card #', 'Title', 'Link', 'Description', 'Checklists', 'Comments', 'Attachments', 'Votes', 'Spent', 'Estimate', 'Created', 'CreatedBy', 'Due', 'Done', 'DoneBy', 'Members', 'Labels']

##### datetimeDone and memberDone
These fields are calculated intercepting when a card was moved to the Done list. While bmccormack's code only checks for this list, I check for cards being moved to any list whose name starts with "Done" (e.g. using lists named "Done Bugfix", "Done New Feature" and so will work).


---

#### Formatting
I tried formatting data in a readable format, suggest changes if you don't like how it is now.

**Comments** are formatted with [date - username] comment, e.g.:

	[2014-01-31 18:38:31 - kathyschultz1] the add-on for exporting to Excel is a good start, but I'm w/all who dream of a report that includes checklists and comments. Thanks Trello warriors!

I added **commentLimit** to limit the number of comments to extract: play with the value (default 100) as per your needs.
There are currently no limits in the number of checklists, checklist items or attachments.

**Attachments** are listed in a similar way, with [filename] (bytes) url, e.g.:

	[chrome.jpg] (62806) https://trello-attachments.s3.amazonaws.com/4d5ea62fd76aa1136000000c/520a29971618ecef3c002181/dc1d95c904a04a6a986b775e55f58bd9/chrome.jpg

**Excel formatting**: after opening the excel you will have to adjust columns widths and formatting. I normally align cells on top and wrap text to have a readable format - see the samples in xlsx.

---

### How to install
This fork is not on the Chrome Web Store, but you can manually install it by following these steps:

1. Download the repository as a zip file
2. Extract zip
3. Go to Chrome Exensions: [chrome://chrome/extensions/](chrome://chrome/extensions/)
4. Click on Developer Mode checkbox
5. Click on Load unpacked extension...
6. Select the folder containing the source files
7. Reload Trello

---

### How to use
1. From a board, click on the board title
2. Click on Share, print, and export...
3. Click on Export Excel

---

### What's next?
I wish to add a configuration dialog to set some options before exporting, e.g. to choose which lists to export, to set the limit on the number of comments to extract etc.
Suggestions are welcome!


---

## Original TrelloExport
### Install Published Version
Grab the original extension from the [Chrome Web Store](https://chrome.google.com/webstore/detail/trelloexport/nhdelomnagopgaealggpgojkhcafhnin?hl=en).
See the original Guthub repository at [https://github.com/llad/export-for-trello](https://github.com/llad/export-for-trello).
