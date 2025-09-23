/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// --- BUILT-IN DICTIONARY (Copied from your Code.gs) ---
const FULL_DICTIONARY = {
  "general": { "good": ["fine", "excellent"], "bad": ["poor", "awful"], /* ... all other dictionaries ... */ },
  // NOTE: To save space, I've omitted the full dictionary. Copy and paste it here from your Code.gs file.
};
const AI_TRACE_PHRASES = [ "certainly, here is", "in conclusion,", "as a large language model," /* ... all other phrases ... */];
// NOTE: Copy and paste the full AI_TRACE_PHRASES array here.


Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Wire up UI event handlers
    document.getElementById('spaceInsertionProb').addEventListener('input', e => document.getElementById('spaceInsertionProbValue').textContent = e.target.value);
    document.getElementById('homoglyphProb').addEventListener('input', e => document.getElementById('homoglyphProbValue').textContent = e.target.value);
    document.getElementById('synonymProb').addEventListener('input', e => document.getElementById('synonymProbValue').textContent = e.target.value);
    document.getElementById('wordInsertionProb').addEventListener('input', e => document.getElementById('wordInsertionProbValue').textContent = e.target.value);

    document.getElementById('findBtn').addEventListener('click', findTraces);
    document.getElementById('deleteBtn').addEventListener('click', deleteTraces);
    document.getElementById('processBtn').addEventListener('click', runProcessor);

    const chkSynonym = document.getElementById('chkSynonym');
    const synonymOptions = document.getElementById('synonym-options');
    synonymOptions.style.display = chkSynonym.checked ? 'block' : 'none';
    chkSynonym.addEventListener('change', () => {
        synonymOptions.style.display = chkSynonym.checked ? 'block' : 'none';
    });
  }
});

function setStatus(message, isError = false) {
    const status = document.getElementById('status');
    status.textContent = message;
    status.style.color = isError ? 'red' : '#6c757d';
}

function disableAllButtons() {
    document.querySelectorAll('button').forEach(btn => btn.disabled = true);
}

function enableAllButtons() {
    document.querySelectorAll('button').forEach(btn => btn.disabled = false);
}

async function findTraces() {
    disableAllButtons();
    setStatus('Finding traces...');
    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            const color = document.getElementById('highlightColor').value;
            let count = 0;

            for (const phrase of AI_TRACE_PHRASES) {
                const searchResults = body.search(phrase, { matchCase: false, matchWholeWord: true });
                context.load(searchResults, 'font');
                await context.sync();
                
                searchResults.items.forEach(item => {
                    item.font.highlightColor = color;
                    count++;
                });
            }
            await context.sync();
            setStatus(count > 0 ? `Highlighted ${count} potential AI traces.` : "No AI traces found.");
        });
    } catch (error) {
        setStatus('Error: ' + error.message, true);
    }
    enableAllButtons();
}

async function deleteTraces() {
    disableAllButtons();
    setStatus('Deleting traces...');
    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            const color = document.getElementById('highlightColor').value;
            let count = 0;

             for (const phrase of AI_TRACE_PHRASES) {
                const searchResults = body.search(phrase, { matchCase: false, matchWholeWord: true });
                context.load(searchResults, 'font');
                await context.sync();
                
                // We check the highlight color to be sure before deleting
                for (const item of searchResults.items) {
                   if (item.font.highlightColor === color) {
                       item.delete();
                       count++;
                   }
                }
            }
            await context.sync();
            setStatus(count > 0 ? `Deleted ${count} highlighted sections.` : "No highlighted sections to delete.");
        });
    } catch (error) {
        setStatus('Error: ' + error.message, true);
    }
    enableAllButtons();
}

async function runProcessor() {
    disableAllButtons();
    setStatus('Processing...');

    try {
        await Word.run(async (context) => {
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load("items/text");
            await context.sync();

            let totalModifications = 0;

            for (const p of paragraphs.items) {
                const originalText = p.text;
                if (!originalText) continue;

                let modificationPlan = buildModificationPlan(originalText);
                if (modificationPlan.length > 0) {
                    applyModificationPlan(p, modificationPlan);
                    totalModifications++;
                }
            }
            await context.sync();
            setStatus(totalModifications > 0 ? "Processing complete!" : "No changes applied.");
        });
    } catch (error) {
         setStatus('Error: ' + error.message, true);
    }
     enableAllButtons();
}

function buildModificationPlan(originalText) {
    const selectedCategories = Array.from(document.querySelectorAll('input[name="synonym-category"]:checked')).map(cb => cb.value);
    const settings = {
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

    let modificationPlan = [];
    let activeSynonymMap = {};
    if (settings.synonym && settings.synonymCategories && settings.synonymCategories.length > 0) {
      settings.synonymCategories.forEach(category => {
        if (FULL_DICTIONARY[category]) {
          Object.assign(activeSynonymMap, FULL_DICTIONARY[category]);
        }
      });
    }

    if (settings.synonym && Object.keys(activeSynonymMap).length > 0) {
      populateSynonymPlan(modificationPlan, originalText, settings.synonymProb, activeSynonymMap);
    }
    if (settings.homoglyph) {
       populateHomoglyphPlan(modificationPlan, originalText, settings.homoglyphProb);
    }
     if (settings.spaceInsertion) {
      populateSpaceInsertionPlan(modificationPlan, originalText, settings.spacePunctuation, settings.spaceInsertionProb);
    }
    if (settings.wordInsertion) {
      populateWordInsertionPlan(modificationPlan, originalText, settings.wordInsertionProb);
    }
    if (settings.spelling) {
      populateSpellingPlan(modificationPlan, originalText, settings.spellingType);
    }
     if (settings.watermark) {
      populateWatermarkPlan(modificationPlan, originalText);
    }
    return modificationPlan;
}


// --- FORMAT-PRESERVING CORE LOGIC for WORD ---
function applyModificationPlan(paragraph, plan) {
  // Sort plan in reverse order by start index. This is crucial for not messing up indices.
  plan.sort((a, b) => b.startIndex - a.startIndex);

  plan.forEach(mod => {
    // For Word, endIndex is exclusive, so we add 1 for replacement.
    const endIndex = (mod.endIndex < mod.startIndex) ? mod.startIndex : mod.endIndex + 1;
    const range = paragraph.getRangeByIndexes(mod.startIndex, endIndex);

    // Differentiate between insertion and replacement.
    if (mod.endIndex < mod.startIndex) { // It's a pure insertion
        range.insertText(mod.newText, Word.InsertLocation.before);
    } else { // It's a replacement or deletion
        range.insertText(mod.newText, Word.InsertLocation.replace);
    }
  });
}

// --- HELPER FUNCTIONS THAT POPULATE THE MODIFICATION PLAN ---
// These functions are pure JavaScript and can be copied directly from your Code.gs file.
// Make sure to copy them all here. I've included one as an example.

function populateSynonymPlan(plan, text, probability, synonymMap) {
  const wordList = Object.keys(synonymMap).join('|');
  const regex = new RegExp(`\\b(${wordList})\\b`, 'gi');
  let match;

  while ((match = regex.exec(text)) !== null) {
    if (Math.random() < probability) {
      const originalWord = match[0];
      const lowerWord = originalWord.toLowerCase();
      const synonyms = synonymMap[lowerWord];
      const chosenSynonym = synonyms[Math.floor(Math.random() * synonyms.length)];

      let replacement = chosenSynonym;
      if (originalWord === originalWord.toUpperCase()) {
        replacement = chosenSynonym.toUpperCase();
      } else if (originalWord.length > 0 && originalWord[0] === originalWord[0].toUpperCase()) {
        replacement = chosenSynonym.charAt(0).toUpperCase() + chosenSynonym.slice(1);
      }

      plan.push({
        startIndex: match.index,
        endIndex: match.index + originalWord.length - 1,
        newText: replacement
      });
    }
  }
}

// ... ADD ALL OTHER populate...Plan FUNCTIONS HERE ...
// populateHomoglyphPlan, populateSpaceInsertionPlan, populateWordInsertionPlan,
// populateSpellingPlan, populateWatermarkPlan
