/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// This is the main entry point for the add-in.
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the necessary APIs.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and initialize UI.
    document.getElementById("processBtn").onclick = () => tryCatch(processText);
    document.getElementById("findBtn").onclick = () => tryCatch(findAndHighlightTraces);
    document.getElementById("deleteBtn").onclick = () => tryCatch(deleteHighlightedTraces);
    
    // UI controls setup
    setupUI();
  }
});


// --- DICTIONARY AND PATTERNS ---
const FULL_DICTIONARY = { /* Omitted for brevity, but this would be the same giant dictionary from Code.gs */ };
const AI_TRACE_PATTERNS = [ /* Omitted for brevity, but this would be the same regex list from Code.gs */ ];

// --- CORE APPLICATION LOGIC ---

async function processText() {
    await Word.run(async (context) => {
        const body = context.document.body;
        const settings = getSettings();
        
        // This is a simplified approach for Word. A full implementation would be more complex.
        // For demonstration, we'll get the whole body text and replace it.
        // NOTE: This will not preserve formatting perfectly like the Google Docs version.
        // A true format-preserving version requires parsing OOXML, which is highly advanced.
        
        const originalText = body.getRange("Whole").load("text");
        await context.sync();

        let modifiedText = originalText.text;

        // Apply transformations
        if (settings.watermark) modifiedText = removeAIWatermarks(modifiedText);
        if (settings.spaceInsertion) modifiedText = insertSpaces(modifiedText, settings.spacePunctuation, settings.spaceInsertionProb);
        if (settings.homoglyph) modifiedText = substituteHomoglyphs(modifiedText, settings.homoglyphProb);
        if (settings.synonym) {
            let activeSynonymMap = {};
            settings.synonymCategories.forEach(category => {
                if (FULL_DICTIONARY[category]) {
                    Object.assign(activeSynonymMap, FULL_DICTIONARY[category]);
                }
            });
            if (Object.keys(activeSynonymMap).length > 0) {
                 modifiedText = replaceSynonyms(modifiedText, settings.synonymProb, activeSynonymMap);
            }
        }
        if (settings.wordInsertion) modifiedText = insertWords(modifiedText, settings.wordInsertionProb);
        if (settings.spelling) modifiedText = varySpelling(modifiedText, settings.spellingType);
        
        if (originalText.text !== modifiedText) {
            body.clear();
            body.insertText(modifiedText, Word.InsertLocation.start);
            updateStatus("Processing complete!");
        } else {
            updateStatus("No changes applied.");
        }
        
        await context.sync();
    });
}

async function findAndHighlightTraces() {
    await Word.run(async (context) => {
        const body = context.document.body;
        const highlightColor = document.getElementById('highlightColor').value;
        const searchPattern = `(?i)(${AI_TRACE_PATTERNS.join('|')})`;

        const searchResults = body.search(searchPattern, { matchCase: false, matchWildcards: true });
        searchResults.load("items");
        await context.sync();

        if (searchResults.items.length > 0) {
            searchResults.items.forEach(range => {
                range.font.highlightColor = highlightColor;
            });
            updateStatus(`Highlighted ${searchResults.items.length} potential AI traces.`);
        } else {
            updateStatus("No AI traces found.");
        }

        await context.sync();
    });
}

async function deleteHighlightedTraces() {
    await Word.run(async (context) => {
        const body = context.document.body;
        const highlightColor = document.getElementById('highlightColor').value;

        // Word API search for formatting is tricky. This is a simplified approach.
        // It searches for any highlighted text and checks if the color matches.
        const searchResults = body.search("*", { matchWildcards: true });
        searchResults.load("items/font");
        await context.sync();
        
        let deletedCount = 0;
        searchResults.items.forEach(range => {
            if (range.font.highlightColor === highlightColor) {
                range.delete();
                deletedCount++;
            }
        });
        
        if (deletedCount > 0) {
             updateStatus(`Deleted ${deletedCount} highlighted sections.`);
        } else {
            updateStatus("No highlighted sections found to delete.");
        }

        await context.sync();
    });
}


// --- UI AND HELPER FUNCTIONS ---

function setupUI() {
    document.getElementById('spaceInsertionProb').addEventListener('input', e => document.getElementById('spaceInsertionProbValue').textContent = e.target.value);
    document.getElementById('homoglyphProb').addEventListener('input', e => document.getElementById('homoglyphProbValue').textContent = e.target.value);
    document.getElementById('synonymProb').addEventListener('input', e => document.getElementById('synonymProbValue').textContent = e.target.value);
    document.getElementById('wordInsertionProb').addEventListener('input', e => document.getElementById('wordInsertionProbValue').textContent = e.target.value);
    
    const chkSynonym = document.getElementById('chkSynonym');
    const synonymOptions = document.getElementById('synonym-options');
    synonymOptions.style.display = chkSynonym.checked ? 'block' : 'none';
    chkSynonym.addEventListener('change', () => {
        synonymOptions.style.display = chkSynonym.checked ? 'block' : 'none';
    });
}

function getSettings() {
    const selectedCategories = Array.from(document.querySelectorAll('input[name="synonym-category"]:checked'))
                                    .map(cb => cb.value);
    return {
        watermark: document.getElementById('chkWatermark').checked,
        spaceInsertion: document.getElementById('chkSpaceInsertion').checked,
        spacePunctuation: document.getElementById('spacePunctuation').value,
        spaceInsertionProb: parseFloat(document.getElementById('spaceInsertionProb').value),
        homoglyph: document.getElementById('chkHomoglyph').checked,
        homoglyphProb: parseFloat(document.getElementById('homoglyphProb').value),
        synonym: document.getElementById('chkSynonym').checked,
        synonymCategories: selectedCategories,
        synonymProb: parseFloat(document.getElementById('synonymProb').value),
        wordInsertion: document.getElementById('chkWordInsertion').checked,
        wordInsertionProb: parseFloat(document.getElementById('wordInsertionProb').value),
        spelling: document.getElementById('chkSpellingVariation').checked,
        spellingType: document.getElementById('spellingVariantType').value
    };
}

// ... (All the text transformation helper functions like removeAIWatermarks, replaceSynonyms, etc. would be pasted here from Code.gs) ...
// NOTE: I've omitted them here for brevity but they are necessary for the code to work.

function updateStatus(message, isError = false) {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = message;
    statusDiv.style.color = isError ? '#dc3545' : '#6c757d';
}

/**
 * A wrapper for asynchronous functions that catches errors and displays them.
 * @param {Function} action An async function to run.
 */
async function tryCatch(action) {
    try {
        updateStatus("Processing...");
        document.querySelectorAll("button").forEach(b => b.disabled = true);
        await action();
    } catch (error) {
        console.error(error);
        updateStatus(`Error: ${error.message}`, true);
    } finally {
        document.querySelectorAll("button").forEach(b => b.disabled = false);
    }
}

