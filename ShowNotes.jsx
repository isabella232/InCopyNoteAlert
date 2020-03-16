//
// ShowNotes.jsx - a script for Adobe InDesign and InCopy 
//
// Version 1.0.3 - 17-March-2020
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
var gLogging = false;

// END OF CONFIGURATION

var gNotesByDocumentKey = {};
var gLogFile = File(Folder.desktop + "/ShowNotes.log");

main();

//--------

function afterOpenHandler(evt) {

    do {

        try {

            if (evt.eventType != Document.AFTER_OPEN) {
                break;
            }

            var document = evt.target;
            if (! (document instanceof Document)) { 
                break;
            }

            if (! document.isValid) {
                break;
            }
        
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
                    try {
                        idleTask.sleep = 0;
                        handleNote();
                    }
                    catch (err) {
                        logError("afterOpenHandler task listener: throws " + err);
                    }
                });
            }

        }
        catch (err) {
            logError("afterOpenHandler: throws " + err);
        }
    }
    while (false);
}

function cleanupNotes() {

    try {

        var seenDocs = {};
        var docCount = app.documents.length;

        for (var docIdx = 0; docIdx < docCount; docIdx++) {
            try {
                var document = app.documents.item(docIdx);
                if (document.isValid) {
                    var documentKey = getDocumentKey(document);
                    seenDocs[documentKey] = docIdx;
                }
            }
            catch (err) { 
                logError("cleanupNotes: loop throws " + err);
            }
        }

        var goneDocs = {};
        for (var documentKey in gNotesByDocumentKey) {
            if (! (documentKey in seenDocs)) {
                goneDocs[documentKey] = true;
            }
        }

        for (var documentKey in goneDocs) {
            delete gNotesByDocumentKey[documentKey];
        }
    }
    catch (err) {    
        logError("cleanupNotes: throws " + err);
    }
}

function detectNotes(document) {

    var hasNotes = false;

    do {

        try {    

            cleanupNotes();
            
            if (! (document instanceof Document)) {
                break;
            }

            if (! document.isValid) {
                break;
            }

            var documentKey = getDocumentKey(document);

            var docEntry = {};
            if (! gNotesByDocumentKey) {
                gNotesByDocumentKey = {};
            }
        
            gNotesByDocumentKey[documentKey] = docEntry;
            
            docEntry.noteEntries = [];
            docEntry.curNoteIdx = 0;

            var storyCount = document.stories.length;
            for (var storyIdx = 0; storyIdx < storyCount; storyIdx++) {
                try {
                    var story = document.stories.item(storyIdx);
                    if (story.isValid) {
                        var storyId = story.id;
                        var noteCount = story.notes.length;
                        for (var noteIdx = 0; noteIdx < noteCount; noteIdx++) {
                            try {
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
                            catch (err) {
                                logError("detectNotes: note loop throws " + err);
                            }
                        }
                    }                    
                }
                catch (err) {
                    logError("detectNotes: story loop throws " + err);
                }
            }
        }
        catch (err) {
            logError("detectNotes: throws " + err);
        }
    }
    while (false);
    
    return hasNotes;
}

function getDocumentKey(document) {

    var documentKey = "invalid:";

    do {
        try {

            if (! (document instanceof Document)) {
                break;
            }

            if (! document.isValid) {
                break;
            }

            if (document.saved) {
                documentKey = document.fullName.absoluteURI;
            }
            else {
                documentKey = "unsaved:" + document.name;
            }

        }
        catch (err) {
            logError("getDocumentKey: throws " + err);
        }
    }
    while (false);

    return documentKey;
}

function gotoCurrentNote() {

    do {

        try {

            var document = app.activeDocument;

            if (! (document instanceof Document)) {
                break;
            }

            if (! document.isValid) {
                break;
            }

            var documentKey = getDocumentKey(document);
            if (! (documentKey in gNotesByDocumentKey)) {
                detectNotes(document);
            }

            var docEntry = gNotesByDocumentKey[documentKey];
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
        catch (err) {            
            logError("gotoCurrentNote: throws " + err);
        }
    }
    while (false);
}

function logError(errorMsg) {

    try {
        if (gLogging) {
            gLogFile.open("a");
            gLogFile.writeln(errorMsg);
            gLogFile.close();
        }        
    }
    catch (err) {
        alert("Critical error: the logError function failed and threw an error: " + err);
    }

}

function main() {

    try {

        //
        // Clear away any old event listeners so we can reinstall them
        //

        app.eventListeners.everyItem().remove();
        app.addEventListener(Document.AFTER_OPEN, afterOpenHandler);

        //
        // Handle the active document if there is one
        //
        if (
            app.documents.length > 0 
        && 
            app.activeDocument 
        && 
            app.activeDocument instanceof Document
        ) {
            afterOpenHandler({ target: app.activeDocument, eventType: Document.AFTER_OPEN});
        }

    } 
    catch (err) {
        logError("main: throws " + err);
    }
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
            logError("showPanel: throws " + err);
        }
    }    
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

