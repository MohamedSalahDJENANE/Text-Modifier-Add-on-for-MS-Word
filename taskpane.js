/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// --- DICTIONARY (Copied from Google Apps Script) ---
const FULL_DICTIONARY = {
  "general": { "good": ["fine", "excellent", "satisfactory", "adequate", "superior"], "bad": ["poor", "awful", "terrible", "inferior"], /* ... and so on ... */ },
  "scientific": { "analysis": ["examination", "investigation", "study"], "data": ["information", "statistics", "figures"], /* ... */ },
  "physics": { "force": ["strength", "power", "energy"], "energy": ["power", "vitality", "force"], /* ... */ },
  "mathematics": { "equation": ["formula", "expression", "calculation"], "variable": ["unknown", "symbol", "placeholder"], /* ... */ },
  "medical": { "patient": ["case", "subject", "sufferer"], "treatment": ["therapy", "remedy", "cure"], /* ... */ },
  "chemistry": { "element": ["substance", "component", "constituent"], "compound": ["mixture", "amalgam", "synthesis"], /* ... */ },
  "literature": { "narrative": ["story", "account", "tale"], "character": ["persona", "figure", "protagonist"], /* ... */ },
  "business": { "strategy": ["plan", "approach", "policy"], "revenue": ["income", "earnings", "turnover"], /* ... */ }
};
// NOTE: For brevity, the full dictionary is truncated here, but you should paste the complete object from your Code.gs file.


// --- AI TRACE PHRASES (Copied from Google Apps Script) ---
const AI_TRACE_PHRASES = [
  "certainly, here is", "in conclusion,", "as a large language model,", "I hope this helps",
  "feel free to ask", "it is important to note", "however, it's also important to consider"
  // NOTE: Paste the complete list from your Code.gs file.
];


Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // --- UI SETUP ---
    // Update slider value displays
    document.getElementById('spaceInsertionProb').addEventListener('input', e => document.getElementById('spaceInsertionProbValue').textContent = e.target.value);
    document.getElementById('homoglyphProb').addEventListener('input', e => document.getElementById('homoglyphProbValue').textContent = e.target.value);
    document.getElementById('synonymProb').addEventListener('input', e => document.getElementById('synonymProbValue').textContent = e.target.value);
    document.getElementById('wordInsertionProb').addEventListener('input', e => document.getElementById('wordInsertionProbValue').textContent = e.target.value);

    // Event Listeners for buttons
    document.getElementById('findBtn').addEventListener('click', findTraces);
    document.getElementById('deleteBtn').addEventListener('click', deleteTraces);
    document.getElementById('processBtn').addEventListener('click', runProcessor);

    // Toggle synonym options visibility
    const chkSynonym = document.getElementById('chkSynonym');
    const synonymOptions = document.getElementById('synonym-options');
    chkSynonym.addEventListener('change', () => {
        synonymOptions.style.display = chkSynonym.checked ? 'block' : 'none';
    });
  }
});

// --- UI HELPERS ---
const statusDiv = document.getElementById('status');
function showStatus(message, isError = false) {
  statusDiv.textContent = message;
  statusDiv.style.color = isError ? 'red' : '#6c757d';
}

function toggleButtons(enable) {
    document.querySelectorAll('button').forEach(btn => btn.disabled = !enable);
}

// --- CORE WORD API FUNCTIONS ---

async function findTraces() {
  try {
    toggleButtons(false);
    showStatus('Finding traces...');
    const color = document.getElementById('highlightColor').value;
    let foundCount = 0;

    await Word.run(async (context) => {
      const body = context.document.body;
      
      // The Word API doesn't support complex regex, so we search for each phrase.
      for (const phrase of AI_TRACE_PHRASES) {
        const searchResults = body.search(phrase, { matchCase: false });
        context.load(searchResults, 'items');
        await context.sync();

        if (searchResults.items.length > 0) {
            foundCount += searchResults.items.length;
            searchResults.items.forEach(item => {
                item.font.highlightColor = color;
            });
        }
      }
      await context.sync();
    });

    showStatus(foundCount > 0 ? `Highlighted ${foundCount} potential AI traces.` : "No AI traces found.");
  } catch (error) {
    showStatus(`Error: ${error.message}`, true);
    console.error(error);
  } finally {
    toggleButtons(true);
  }
}

async function deleteTraces() {
    try {
        toggleButtons(false);
        showStatus('Deleting traces...');
        const color = document.getElementById('highlightColor').value;
        let deletedCount = 0;

        await Word.run(async (context) => {
            const body = context.document.body;
            
            // Re-find the phrases to delete them. A direct "find by format" is complex.
            for (const phrase of AI_TRACE_PHRASES) {
                const searchResults = body.search(phrase, { matchCase: false });
                context.load(searchResults, 'items');
                await context.sync();

                if (searchResults.items.length > 0) {
                    searchResults.items.forEach(item => {
                        // We check the highlight color to be sure, although it's not foolproof.
                        // This is a limitation compared to the original script's approach.
                        item.font.load('highlightColor');
                    });
                    await context.sync();

                    searchResults.items.forEach(item => {
                       if (item.font.highlightColor && item.font.highlightColor.toLowerCase() === color.toLowerCase()) {
                           item.insertText('', Word.InsertLocation.replace);
                           deletedCount++;
                       }
                    });
                }
            }
            await context.sync();
        });

        showStatus(deletedCount > 0 ? `Deleted ${deletedCount} highlighted phrases.` : "No highlighted phrases found to delete.");
    } catch (error) {
        showStatus(`Error: ${error.message}`, true);
        console.error(error);
    } finally {
        toggleButtons(true);
    }
}


async function runProcessor() {
  try {
    toggleButtons(false);
    showStatus('Processing...');
    
    // 1. Get Settings from UI
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

    let totalModifications = 0;

    await Word.run(async (context) => {
      // 2. Build active synonym map
      let activeSynonymMap = {};
      if (settings.synonym && settings.synonymCategories.length > 0) {
        settings.synonymCategories.forEach(category => {
          if (FULL_DICTIONARY[category]) {
            Object.assign(activeSynonymMap, FULL_DICTIONARY[category]);
          }
        });
      }

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();

      // 3. Iterate through paragraphs to build and apply modification plans
      for (let i = 0; i < paragraphs.items.length; i++) {
        const p = paragraphs.items[i];
        const originalText = p.text;
        if (!originalText.trim()) continue;

        // Create a modification plan for this specific paragraph's text.
        let modificationPlan = [];

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

        if (modificationPlan.length > 0) {
            totalModifications++;
            // Apply the plan for this paragraph
            // Sort plan in reverse order by start index. This is crucial.
            modificationPlan.sort((a, b) => b.startIndex - a.startIndex);

            modificationPlan.forEach(mod => {
                // The Word API endIndex is exclusive, so we might need +1 depending on logic
                // For replacement, we use startIndex and endIndex. For insertion, endIndex < startIndex.
                const start = mod.startIndex;
                const end = (mod.endIndex < start) ? start : mod.endIndex + 1;
                
                const range = p.getRangeByIndexes(start, end);
                range.insertText(mod.newText, Word.InsertLocation.replace);
            });
        }
      } // end for loop
      await context.sync();
    });

    showStatus(totalModifications > 0 ? "Processing complete!" : "No changes applied.");
  } catch (error) {
    showStatus(`Error: ${error.message}`, true);
    console.error(error);
  } finally {
    toggleButtons(true);
  }
}


// --- MODIFICATION PLAN BUILDERS (Copied from Google Apps Script, no changes needed) ---

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

function populateHomoglyphPlan(plan, text, probability) {
  const homoglyphs = { 'a': 'а', 'e': 'е', 'o': 'о', 'c': 'с', 'i': 'і', 'p': 'р', 's': 'ѕ', 'x': 'х' };
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (homoglyphs[char] && Math.random() < probability) {
      plan.push({
        startIndex: i,
        endIndex: i,
        newText: homoglyphs[char]
      });
    }
  }
}

function populateSpaceInsertionPlan(plan, text, punctuationStr, probability) {
    if (!punctuationStr) return;
    const escapedPunctuation = punctuationStr.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
    const regex = new RegExp(`(\\S)([${escapedPunctuation}])`, 'g');
    let match;
    while ((match = regex.exec(text)) !== null) {
        if (Math.random() < probability) {
            plan.push({
                startIndex: match.index + 1,
                endIndex: match.index, // Signifies insertion
                newText: " "
            });
        }
    }
}

function populateWordInsertionPlan(plan, text, probability) {
    const commonWords = ["actually", "basically", "literally", "really", "just", "sort of", "kind of", "well", "essentially"];
    const regex = /\b(\w{3,})\b/g; 
    let match;
    while ((match = regex.exec(text)) !== null) {
        if (Math.random() < probability) {
            const word = match[0];
            const insertedWord = " " + commonWords[Math.floor(Math.random() * commonWords.length)];
            plan.push({
                startIndex: match.index + word.length,
                endIndex: match.index + word.length - 1, // Insertion
                newText: insertedWord
            });
        }
    }
}

function populateSpellingPlan(plan, text, direction) {
    const usToUk = {"analyze":"analyse","behavior":"behaviour","center":"centre","color":"colour","defense":"defence","favorite":"favourite"};
    const ukToUs = Object.keys(usToUk).reduce((obj, key) => { obj[usToUk[key]] = key; return obj; }, {});
    const map = direction === "usToUk" ? usToUk : (direction === "ukToUs" ? ukToUs : (Math.random() < 0.5 ? usToUk : ukToUs));
    
    const wordList = Object.keys(map).join('|');
    const regex = new RegExp(`\\b(${wordList})\\b`, 'gi');
    let match;

    while ((match = regex.exec(text)) !== null) {
        const originalWord = match[0];
        const lowerWord = originalWord.toLowerCase();
        const replacement = map[lowerWord];

        let finalReplacement = replacement;
        if (originalWord === originalWord.toUpperCase()) {
            finalReplacement = replacement.toUpperCase();
        } else if (originalWord.length > 0 && originalWord[0] === originalWord[0].toUpperCase()) {
            finalReplacement = replacement.charAt(0).toUpperCase() + replacement.slice(1);
        }

        plan.push({
            startIndex: match.index,
            endIndex: match.index + originalWord.length - 1,
            newText: finalReplacement
        });
    }
}

function populateWatermarkPlan(plan, text) {
    const invisibleRegex = /[\u200B-\u200D\uFEFF]/g;
    let match;
    while ((match = invisibleRegex.exec(text)) !== null) {
        plan.push({ startIndex: match.index, endIndex: match.index, newText: "" });
    }

    const multiSpaceRegex = /\s{2,}/g;
     while ((match = multiSpaceRegex.exec(text)) !== null) {
        plan.push({ startIndex: match.index, endIndex: match.index + match[0].length - 1, newText: " " });
    }
}
