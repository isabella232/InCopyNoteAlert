//
// ShowNotes.jsx - a script for Adobe InDesign and InCopy 
//
// Version 1.02 - 7-March-2020
//
// by Kris Coppieters 
// kris@rorohiko.com
// https://www.linkedin.com/in/kristiaan/
// ----------------
//
// About Rorohiko:
//
// Rorohiko specialises in making printing, publishing and web workflows more efficient.
//
// This script is a free sample of the custom solutions we create for our customers.
//
// If your workflow is hampered by boring or repetitive tasks, inquire at
//
//   sales@rorohiko.com
//
// The scripts we write for our customers repay for themselves within weeks or 
// months.
//
// ---------------
//
// About this script:
//
// This script will cause InDesign or InCopy to display a warning dialog 
// each time you open a document that has editorial notes embedded in it
// (as created with panel accessed via Window - Editorial - Notes menu). 
//
// For licensing and copyright information: see end of the script
//
// For installation info and documentation: visit 
//
// https://github.com/zwettemaan/InCopyNoteAlert/wiki
//
// To tweak this script: carefully make sensible adjustments between the two lines
// 'CONFIFURATION' - 'END OF CONFIGURATION' below, and make sure to restart InDesign
// or InCopy after making changes.
// 

#targetengine com.rorohiko.CreativeProWeek.ShowNotes

// CONFIGURATION:

var gShowAlert = true;
var gGotoFirstNote = true;
var gZoomPercentage = 250;
var gShowPanel = true;
var gMillisecondsToWaitForIdle = 1000;

// END OF CONFIGURATION

var gNotesByDocumentURL = {};

function cleanupNotes() {

    try {
        var seenDocs = {};
        var docCount = app.documents.length;
        for (var docIdx = 0; docIdx < docCount; docIdx++) {
            var document = app.documents.item(docIdx);
            var url = getURL(document);
            seenDocs[url] = docIdx;            
        }

        var goneDocs = {};
        for (var url in gNotesByDocumentURL) {
            if (! (url in seenDocs)) {
                goneDocs[url] = true;
            }
        }

        for (var url in goneDocs) {
            delete gNotesByDocumentURL[url];
        }
    }
    catch (err) {    
    }

}

function gotoCurrentNote() {

    do {

        if (! (app.activeDocument instanceof Document)) {
            break;
        }

        var document = app.activeDocument;
        var url = getURL(document);
        if (! (url in gNotesByDocumentURL)) {
            detectNotes(document);
        }

        var docEntry = gNotesByDocumentURL[url];
        if (! docEntry) {
            break;
        }

        if 
        (
            docEntry.curNoteIdx < 0 
        || 
            docEntry.curNoteIdx >= docEntry.noteEntries.length
        ) {
            docEntry.curNoteIdx = 0;
        }

        if (docEntry.curNoteIdx >= docEntry.noteEntries.length) {
            break;
        }

        var noteEntry = docEntry.noteEntries[docEntry.curNoteIdx];
        if (! noteEntry) {
            break;
        }

        var story = document.stories.itemByID(noteEntry.storyId);
        if (! story || ! story.isValid) {
            break;
        }

        var note = story.notes.itemByID(noteEntry.noteId);
        if (! note || ! note.isValid) {
            break;
        }

        app.selection = story.characters.item(note.storyOffset.index);  
        
        story.characters.item(note.storyOffset.index).showText();
        
        app.activeWindow.zoomPercentage = gZoomPercentage;
        
        if (gShowPanel) {
            showPanel("Notes");
        }
    }
    while (false);
}

function getURL(document) {
    var url;
    if (document.saved) {
        url = document.fullName.absoluteURI;
    }
    else {
        url = document.name;
    }

    return url;
}

function detectNotes(document) {

    var hasNotes = false;

    do {
        try {    

            cleanupNotes();
            
            if (! (document instanceof Document)) {
                break;
            }

            var url = getURL(document);

            var docEntry = {};
            if (! gNotesByDocumentURL) {
                gNotesByDocumentURL = {};
            }
        
            gNotesByDocumentURL[url] = docEntry;
            
            docEntry.noteEntries = [];
            docEntry.curNoteIdx = 0;

            var storyCount = document.stories.length;
            for (var storyIdx = 0; storyIdx < storyCount; storyIdx++) {
                var story = document.stories.item(storyIdx);
                if (story.isValid) {
                    var storyId = story.id;
                    var noteCount = story.notes.length;
                    for (var noteIdx = 0; noteIdx < noteCount; noteIdx++) {
                        var note = story.notes.item(noteIdx);
                        if (note.isValid) {
                            var noteId = note.id;
                            var noteEntry = {
                                storyId: storyId,
                                noteId: noteId
                            };
                            hasNotes = true;
                            docEntry.noteEntries.push(noteEntry);
                        }
                    }
                }
            }
        }
        catch (err) {
        }
    }
    while (false);
    
    return hasNotes;
}

function afterOpenHandler(evt) {

    do {

        try {

            if (evt.eventType != Document.AFTER_OPEN) {
                break;
            }

            if (! evt.target instanceof Document) { 
                break;
            }

            var document = evt.target;
            var hasNotes = detectNotes(document);

            var handleNote = function() {
                alert("The 'ShowNotes.jsx' startup script reports:\nThere are Notes in this document.");
                if (gGotoFirstNote) {
                    gotoCurrentNote();
                }
            };
            
            if (hasNotes && gShowAlert) {
              var idleTask = app.idleTasks.add({
                sleep: gMillisecondsToWaitForIdle
              });    
          
              var noteTaskEventListener = idleTask.addEventListener(IdleTask.ON_IDLE, function() {
                idleTask.sleep = 0;
                handleNote();
              });
            }

        }
        catch (err) {
        }
    }
    while (false);
}

function showPanel(panelName) {

    var panelCount = app.panels.length;
    for (var idx = 0; idx < panelCount; idx++) {
        try {
            var panel = app.panels.item(idx);
            if (panel.name == panelName) {
                panel.visible = true;
            }
        }
        catch (err) {
        }
    }
    
}

try {

    //
    // Clear away any old event listeners so we can reinstall them
    //

    app.eventListeners.everyItem().remove();
    app.addEventListener(Document.AFTER_OPEN, afterOpenHandler);

    //
    // Handle the active document if there is one
    //
    if (app.documents.length > 0 && app.activeDocument && app.activeDocument instanceof Document) {
        afterOpenHandler({ target: app.activeDocument, eventType: Document.AFTER_OPEN});
    }

} 
catch (err) {
}

/*************************************************************

ShowNotes.jsx

MIT License

Copyright (c) 2020 2018-2020 Rorohiko Ltd. - Kris Coppieters - kris@rorohiko.com

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

==============================================
*/

