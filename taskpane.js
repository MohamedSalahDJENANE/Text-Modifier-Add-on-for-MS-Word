/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// This is the core function that ensures the Office environment is ready before running any code.
Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    // Wait for the DOM to be fully loaded before attaching event handlers
    document.addEventListener("DOMContentLoaded", function() {
      // Attach event listeners to all UI elements
      document.getElementById("processBtn").addEventListener("click", processDocument);
      document.getElementById("highlightBtn").addEventListener("click", () => findAndHighlightTraces(true));
      document.getElementById("deleteBtn").addEventListener("click", () => findAndHighlightTraces(false));

      // Slider value displays
      document.getElementById('spaceInsertionProb').addEventListener('input', e => document.getElementById('spaceInsertionProbValue').textContent = e.target.value);
      document.getElementById('homoglyphProb').addEventListener('input', e => document.getElementById('homoglyphProbValue').textContent = e.target.value);
      document.getElementById('synonymProb').addEventListener('input', e => document.getElementById('synonymProbValue').textContent = e.target.value);
      document.getElementById('wordInsertionProb').addEventListener('input', e => document.getElementById('wordInsertionProbValue').textContent = e.target.value);

      // Toggle synonym options visibility
      const chkSynonym = document.getElementById('chkSynonym');
      const synonymOptions = document.getElementById('synonym-options');
      synonymOptions.style.display = chkSynonym.checked ? 'block' : 'none';
      chkSynonym.addEventListener('change', () => {
          synonymOptions.style.display = chkSynonym.checked ? 'block' : 'none';
      });
    });
  }
});

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

async function processDocument() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items/text"); // Load paragraph text
      await context.sync();

      const settings = getSettings();
      let activeSynonymMap = {};
      if (settings.synonym && settings.synonymCategories.length > 0) {
        settings.synonymCategories.forEach(category => {
          if (FULL_DICTIONARY[category]) {
            Object.assign(activeSynonymMap, FULL_DICTIONARY[category]);
          }
        });
      }
      
      for (const paragraph of paragraphs.items) {
        // We must process paragraph by paragraph for Word.js
        const originalText = paragraph.text;
        let modifiedText = originalText;

        // Apply modifications
        if (settings.watermark) modifiedText = modifiedText.replace(/[\u200B-\u200D\uFEFF]/g, '').replace(/\s{2,}/g, ' ');
        if (settings.spaceInsertion) modifiedText = insertSpaces(modifiedText, settings.spacePunctuation, settings.spaceInsertionProb);
        if (settings.homoglyph) modifiedText = substituteHomoglyphs(modifiedText, settings.homoglyphProb);
        if (settings.synonym && Object.keys(activeSynonymMap).length > 0) modifiedText = replaceSynonyms(modifiedText, settings.synonymProb, activeSynonymMap);
        if (settings.wordInsertion) modifiedText = insertWords(modifiedText, settings.wordInsertionProb);
        if (settings.spelling) modifiedText = varySpelling(modifiedText, settings.spellingType);
        
        if(originalText !== modifiedText) {
            paragraph.insertText(modifiedText, Word.InsertLocation.replace);
        }
      }

      await context.sync();
      updateStatus("Processing complete!");
    });
  } catch (error) {
    console.error(error);
    updateStatus("Error: " + error.message);
  }
}

async function findAndHighlightTraces(highlight) {
    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            const highlightColor = document.getElementById('highlightColor').value;
            let count = 0;

            if (highlight) {
                // Using the Word API's built-in regex search is much more efficient.
                const searchPattern = `(${AI_TRACE_PATTERNS.join('|')})`;
                const searchResults = body.search(searchPattern, { matchCase: false, matchWildCards: true });
                searchResults.load("items");
                await context.sync();

                searchResults.items.forEach(item => {
                    item.font.highlightColor = highlightColor;
                    count++;
                });
                updateStatus(`Highlighted ${count} traces.`);
            } else {
                // To remove, we search for any text with the specific highlight color.
                // This is a workaround since Word.js doesn't directly support finding highlights.
                // We'll have to iterate.
                const paragraphs = context.document.body.paragraphs;
                paragraphs.load("items/font");
                await context.sync();

                for (const para of paragraphs.items) {
                    const searchResults = para.search("*", {matchWildCards: true});
                    searchResults.load("items/font");
                    await context.sync();

                    searchResults.items.forEach(range => {
                        if (range.font.highlightColor === highlightColor) {
                            range.delete();
                            count++;
                        }
                    });
                }
                updateStatus(`Deleted ${count} highlighted sections.`);
            }
            await context.sync();
        });
    } catch (error) {
        console.error(error);
        updateStatus("Error: " + error.message);
    }
}


function updateStatus(message) {
  const statusDiv = document.getElementById('status');
  statusDiv.textContent = message;
}

// --- HELPER FUNCTIONS ---

function insertSpaces(text, punctuationStr, probability) {
  if (!punctuationStr || probability === 0) return text;
  const punctuationChars = new RegExp(`([\\S])([${punctuationStr.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&')}])`, 'g');
  return text.replace(punctuationChars, (match, charBefore, punc) => {
    if (Math.random() < probability) {
      return `${charBefore} ${punc}`;
    }
    return match;
  });
}

function substituteHomoglyphs(text, probability) {
  if (probability === 0) return text;
  const homoglyphs = { 'a': 'а', 'e': 'е', 'o': 'о', 'c': 'с', 'i': 'і', 'p': 'р', 's': 'ѕ', 'x': 'х', 'A': 'А', 'E': 'Е', 'O': 'О', 'C': 'С', 'I': 'І', 'P': 'Р', 'S': 'Ѕ', 'X': 'Х' };
  return text.split('').map(char => (homoglyphs[char] && Math.random() < probability) ? homoglyphs[char] : char).join('');
}

function replaceSynonyms(text, probability, synonymMap) {
  const wordList = Object.keys(synonymMap).join('|');
  const regex = new RegExp(`\\b(${wordList})\\b`, 'gi');
  return text.replace(regex, (originalWord) => {
    if (Math.random() < probability) {
      const lowerWord = originalWord.toLowerCase();
      const synonyms = synonymMap[lowerWord];
      const chosenSynonym = synonyms[Math.floor(Math.random() * synonyms.length)];

      if (originalWord === originalWord.toUpperCase()) return chosenSynonym.toUpperCase();
      if (originalWord[0] === originalWord[0].toUpperCase()) return chosenSynonym.charAt(0).toUpperCase() + chosenSynonym.slice(1);
      return chosenSynonym;
    }
    return originalWord;
  });
}

function insertWords(text, probability) {
  if (probability === 0) return text;
  const commonWords = ["actually", "basically", "literally", "really", "just", "sort of", "kind of", "well", "essentially"];
  return text.replace(/\b(\w{3,})\b/g, (word) => {
    if (Math.random() < probability) {
      return word + " " + commonWords[Math.floor(Math.random() * commonWords.length)];
    }
    return word;
  });
}

function varySpelling(text, direction) {
  const usToUk = {"analyze":"analyse","behavior":"behaviour","center":"centre","color":"colour","defense":"defence","favorite":"favourite","flavor":"flavour","gray":"grey","humor":"humour","labor":"labour","license":"licence","neighbor":"neighbour","organize":"organise","realize":"realise","recognize":"recognise","theater":"theatre","traveled":"travelled"};
  const ukToUs = Object.keys(usToUk).reduce((obj, key) => { obj[usToUk[key]] = key; return obj; }, {});
  const map = direction === "usToUk" ? usToUk : (direction === "ukToUs" ? ukToUs : (Math.random() < 0.5 ? usToUk : ukToUs));
  const wordList = Object.keys(map).join('|');
  const regex = new RegExp(`\\b(${wordList})\\b`, 'gi');
  
  return text.replace(regex, (originalWord) => {
    const lowerWord = originalWord.toLowerCase();
    const replacement = map[lowerWord];
    if (originalWord === originalWord.toUpperCase()) return replacement.toUpperCase();
    if (originalWord[0] === originalWord[0].toUpperCase()) return replacement.charAt(0).toUpperCase() + replacement.slice(1);
    return replacement;
  });
}

// --- BUILT-IN DICTIONARIES & PATTERNS ---
const AI_TRACE_PATTERNS = [
  "certainly, here('s| is)", "of course, here('s| is)", "sure, here('s| is)",
  "here('s| is) (a|an|the|your|what i found|what i came up with|how you can|how you could)",
  "here is a brief", "here is a summary", "here is an outline",
  "here is an introduction for you", "here's an introduction for you", 
  "here is a summary for you", "here's a summary for you",
  "certainly, i can help with that",
  "let’s break it down", "let me explain",
  "to begin with,?", "firstly,?", "first of all,?",
  "in conclusion,", "in summary,", "to summarize,", "in sum,", "to sum up,", "overall,",
  "put simply,", "in short,", "all in all,", "the key takeaway is",
  "the main idea is", "ultimately,", "the bottom line is",
  "(i )?hope this (helps|was helpful|is helpful|information helps|is useful)",
  "let me know if you have any (other|further) questions",
  "let me know if you need anything else",
  "feel free to (ask|reach out)",
  "please don’t hesitate to ask",
  "as a large language model,", "as an ai language model,", "as an ai,",
  "as an artificial intelligence,", "i am an ai,", "i’m an ai,",
  "i (cannot|can't|am not able to|am unable to)",
  "i do not have the ability to", "i don’t have personal opinions",
  "i don’t have beliefs", "i do not have beliefs",
  "i don’t have personal experiences", "i lack personal experiences",
  "my knowledge cutoff is", "my training data only goes up to",
  "my knowledge is current up to",
  "it is important to note", "it should be noted", "it’s worth noting that",
  "it is also important to note", "please note that",
  "however, it’s also important to consider",
  "it’s important to remember that",
  "keep in mind that", "one thing to keep in mind",
  "additionally,", "furthermore,", "moreover,",
  "let’s go step by step", "let’s go through this",
  "to put it another way", "in other words,",
  "let’s break this down", "to clarify,",
  "this means that", "what this implies is",
  "you could try", "you might consider",
  "one approach is", "another option is",
  "a common way to do this is", "a possible solution is",
  "an alternative is", "the recommended way is"
];

const FULL_DICTIONARY = {
  "general": {
    "good": ["fine", "excellent", "satisfactory", "adequate", "superior", "acceptable", "proficient", "competent", "virtuous", "favorable", "splendid"],
    "bad": ["poor", "awful", "terrible", "inferior", "substandard", "deficient", "inadequate", "deplorable", "unsatisfactory", "dreadful"],
    "important": ["significant", "crucial", "vital", "essential", "paramount", "pivotal", "critical", "consequential", "momentous", "indispensable"],
    "happy": ["joyful", "pleased", "content", "delighted", "elated", "cheerful", "jubilant", "gleeful", "ecstatic", "merry"],
    "sad": ["unhappy", "sorrowful", "dejected", "melancholy", "depressed", "glum", "despondent", "disheartened", "mournful"],
    "big": ["large", "huge", "massive", "enormous", "vast", "immense", "colossal", "gigantic", "substantial", "monumental"],
    "small": ["tiny", "little", "minuscule", "compact", "minute", "diminutive", "microscopic", "slight", "petty"],
    "fast": ["quick", "rapid", "swift", "hasty", "speedy", "brisk", "expeditious", "nimble", "prompt"],
    "slow": ["leisurely", "gradual", "unhurried", "plodding", "deliberate", "measured", "lingering", "sluggish"],
    "help": ["assist", "support", "aid", "facilitate", "serve", "abet", "succor", "encourage", "sustain"],
    "use": ["utilize", "employ", "operate", "apply", "leverage", "harness", "exploit", "handle", "manage"],
    "make": ["create", "produce", "construct", "build", "fabricate", "manufacture", "assemble", "form", "generate"],
    "explain": ["clarify", "describe", "define", "illustrate", "elucidate", "interpret", "expound", "explicate", "justify"],
    "show": ["display", "reveal", "present", "demonstrate", "exhibit", "indicate", "manifest", "disclose", "convey"],
    "interesting": ["fascinating", "engaging", "intriguing", "compelling", "gripping", "captivating", "thought-provoking", "stimulating"],
    "idea": ["concept", "notion", "thought", "suggestion", "insight", "perception", "impression", "conception"],
    "problem": ["issue", "challenge", "obstacle", "difficulty", "complication", "predicament", "dilemma", "quandary"],
    "different": ["distinct", "dissimilar", "unlike", "varied", "diverse", "contrasting", "disparate", "unique"],
    "begin": ["start", "commence", "initiate", "launch", "originate", "institute", "found", "inaugurate"],
    "end": ["finish", "conclude", "terminate", "cease", "complete", "finalize", "culminate", "halt"],
    "change": ["alter", "modify", "transform", "vary", "adjust", "revise", "amend", "convert"],
    "get": ["obtain", "acquire", "receive", "procure", "attain", "secure", "derive", "gain"]
  },
  "scientific": {
    "analysis": ["examination", "investigation", "study", "scrutiny", "evaluation", "assessment", "inquiry", "breakdown", "dissection"],
    "data": ["information", "statistics", "figures", "facts", "measurements", "observations", "findings", "evidence", "input"],
    "evidence": ["proof", "confirmation", "substantiation", "validation", "grounds", "testimony", "corroboration", "indication", "attestation"],
    "hypothesis": ["theory", "supposition", "premise", "assumption", "postulate", "conjecture", "proposition", "thesis"],
    "method": ["procedure", "technique", "approach", "system", "process", "methodology", "protocol", "regimen"],
    "result": ["outcome", "consequence", "finding", "conclusion", "effect", "upshot", "determination", "product", "yield"],
    "research": ["investigation", "study", "experimentation", "exploration", "inquiry", "analysis", "examination", "fact-finding"],
    "theory": ["principle", "concept", "doctrine", "hypothesis", "framework", "ideology", "model", "supposition"],
    "variable": ["factor", "element", "condition", "characteristic", "property", "determinant", "parameter", "unknown"],
    "experiment": ["test", "trial", "procedure", "investigation", "demonstration", "protocol", "study", "probe"],
    "observation": ["scrutiny", "monitoring", "surveillance", "finding", "remark", "notation", "examination", "perception"],
    "conclusion": ["deduction", "inference", "resolution", "summary", "verdict", "determination", "judgement", "culmination"],
    "significant": ["meaningful", "notable", "consequential", "substantial", "critical", "momentous", "statistically relevant"],
    "quantitative": ["numerical", "measurable", "statistical", "computable", "quantifiable", "numeric"],
    "qualitative": ["descriptive", "observational", "interpretive", "subjective", "non-numerical", "conceptual"],
    "component": ["element", "part", "constituent", "ingredient", "factor", "module", "unit"],
    "factor": ["element", "component", "determinant", "influence", "variable", "consideration"],
    "feature": ["characteristic", "quality", "property", "attribute", "trait", "aspect", "hallmark"],
    "framework": ["structure", "scaffolding", "model", "system", "schema", "architecture", "skeleton"],
    "model": ["representation", "paradigm", "prototype", "framework", "simulation", "construct", "depiction"],
    "paradigm": ["model", "pattern", "example", "framework", "archetype", "standard", "exemplar"],
    "phenomenon": ["occurrence", "event", "fact", "happening", "anomaly", "process", "manifestation"],
    "principle": ["law", "rule", "standard", "axiom", "doctrine", "tenet", "fundamental"],
    "property": ["characteristic", "attribute", "quality", "feature", "trait", "peculiarity"],
    "relationship": ["connection", "correlation", "association", "link", "correspondence", "interrelation", "nexus"],
    "simulate": ["imitate", "replicate", "model", "reproduce", "emulate", "mimic"],
    "source": ["origin", "derivation", "root", "basis", "foundation", "provenance"],
    "specimen": ["sample", "example", "instance", "prototype", "model", "unit"],
    "structure": ["arrangement", "formation", "configuration", "composition", "makeup", "organization", "anatomy"],
    "system": ["arrangement", "network", "structure", "organization", "complex", "methodology", "entity"],
    "technique": ["method", "procedure", "approach", "way", "means", "tactic"],
    "validate": ["confirm", "verify", "substantiate", "corroborate", "authenticate", "certify"],
    "yield": ["produce", "generate", "provide", "result in", "give", "afford", "supply"]
  },
  "physics": {
    "force": ["strength", "power", "energy", "pressure", "impact", "momentum", "impetus", "compulsion", "drive"],
    "energy": ["power", "vitality", "force", "vigor", "potential", "capacity", "work", "dynamism"],
    "velocity": ["speed", "rate", "pace", "momentum", "celerity", "rapidity", "tempo"],
    "acceleration": ["hastening", "quickening", "increase in speed", "rate of velocity change", "speedup"],
    "mass": ["weight", "bulk", "magnitude", "density", "substance", "inertia", "quantity of matter"],
    "field": ["domain", "area", "region", "spectrum", "influence", "force field", "expanse"],
    "particle": ["speck", "grain", "fragment", "molecule", "atom", "corpuscle", "subatomic particle", "elementary particle"],
    "quantum": ["unit", "measure", "quantity", "increment", "packet", "discrete unit", "portion"],
    "relativity": ["interconnection", "correspondence", "proportionality", "interdependence", "comparative nature"],
    "wave": ["oscillation", "ripple", "undulation", "vibration", "fluctuation", "pulse", "surge"],
    "spectrum": ["range", "band", "continuum", "distribution", "array", "scale"],
    "gravity": ["gravitation", "attraction", "pull", "downward force", "gravitational force"],
    "magnetism": ["attraction", "polarity", "magnetic force", "magnetic field"],
    "radiation": ["emission", "rays", "energy waves", "radioactivity", "flux"],
    "theorem": ["law", "principle", "postulate", "axiom", "proposition"],
    "dynamics": ["mechanics", "kinetics", "motion studies", "forces in motion"],
    "kinetics": ["dynamics", "motion analysis", "study of motion", "movement science"],
    "momentum": ["impetus", "driving force", "velocity", "inertia", "product of mass and velocity"],
    "oscillation": ["vibration", "fluctuation", "swing", "wave", "undulation", "pulsation"],
    "resonance": ["vibration", "reverberation", "amplification", "sympathetic vibration"],
    "statics": ["equilibrium", "balance", "study of forces in equilibrium", "stability"],
    "thermodynamics": ["heat transfer", "energy conversion", "study of heat", "heat dynamics"],
    "torque": ["turning force", "rotation", "twist", "moment", "rotational force"],
    "vector": ["quantity", "course", "heading", "directed quantity", "directional value"]
  },
  "mathematics": {
    "equation": ["formula", "expression", "calculation", "statement", "identity", "proposition", "mathematical sentence"],
    "variable": ["unknown", "symbol", "placeholder", "quantity", "parameter", "indeterminate", "argument"],
    "constant": ["fixed value", "parameter", "given", "unchanging number", "invariant", "scalar", "fixed quantity"],
    "function": ["operation", "relation", "transformation", "mapping", "correspondence", "dependency"],
    "integral": ["summation", "antiderivative", "total", "aggregation", "calculus of areas", "integration"],
    "derivative": ["rate of change", "gradient", "slope", "differential", "fluxion", "rate"],
    "theorem": ["principle", "law", "postulate", "proposition", "axiom", "lemma", "corollary"],
    "proof": ["verification", "validation", "demonstration", "evidence", "corroboration", "argument", "justification"],
    "matrix": ["array", "grid", "table", "vector", "tensor", "rectangular array", "arrangement"],
    "algorithm": ["procedure", "process", "method", "formula", "set of rules", "routine", "computation"],
    "probability": ["likelihood", "chance", "odds", "prospect", "statistical chance", "expectancy"],
    "axiom": ["postulate", "precept", "truism", "fundamental principle", "given"],
    "calculus": ["infinitesimal calculus", "analysis", "differentiation", "integration"],
    "coefficient": ["factor", "multiplier", "numerical constant", "scalar multiplier"],
    "corollary": ["consequence", "result", "deduction", "inference", "natural result"],
    "denominator": ["divisor", "lower part of a fraction", "base of a fraction"],
    "exponent": ["power", "index", "logarithm", "degree"],
    "fraction": ["ratio", "quotient", "part", "portion", "percentage"],
    "geometry": ["study of shapes", "spatial mathematics", "topology", "study of space"],
    "graph": ["diagram", "chart", "plot", "network", "visual representation"],
    "lemma": ["subsidiary proposition", "preliminary theorem", "assumption", "helping theorem"],
    "logarithm": ["exponent", "power", "index", "log"],
    "numerator": ["dividend", "upper part of a fraction", "top of a fraction"],
    "parameter": ["variable", "limit", "framework", "guideline", "characteristic"],
    "permutation": ["arrangement", "combination", "ordering", "sequence", "rearrangement"],
    "postulate": ["axiom", "premise", "assumption", "supposition", "fundamental"],
    "quotient": ["result of division", "ratio", "fraction", "division result"],
    "radius": ["semidiameter", "spoke", "distance from center to edge"],
    "ratio": ["proportion", "relationship", "fraction", "quotient", "comparison"],
    "scalar": ["magnitude", "constant", "single number", "non-vector quantity"],
    "sequence": ["series", "progression", "succession", "chain", "string"],
    "series": ["sequence", "progression", "succession", "sum", "summation"],
    "set": ["collection", "group", "aggregate", "class", "ensemble"],
    "statistics": ["data analysis", "numerical facts", "quantitative data", "study of data"],
    "topology": ["study of geometric properties", "spatial analysis", "rubber-sheet geometry"],
    "vector": ["directed quantity", "course", "heading", "magnitude with direction"]
  },
  "medical": {
    "patient": ["case", "subject", "sufferer", "invalid", "recipient of care", "convalescent"],
    "treatment": ["therapy", "remedy", "cure", "protocol", "regimen", "intervention", "medication"],
    "diagnosis": ["assessment", "evaluation", "conclusion", "opinion", "identification", "prognosis", "determination"],
    "symptom": ["indication", "sign", "manifestation", "expression", "warning", "indicator", "clue"],
    "disease": ["illness", "sickness", "disorder", "ailment", "condition", "malady", "pathology"],
    "medication": ["drug", "pharmaceutical", "remedy", "prescription", "treatment", "medicament", "formula"],
    "clinical": ["medical", "therapeutic", "hospital", "patient-oriented", "observational", "bedside"],
    "syndrome": ["condition", "disorder", "complex of symptoms", "pattern", "set of signs"],
    "acute": ["severe", "sharp", "intense", "critical", "sudden onset"],
    "chronic": ["persistent", "long-lasting", "prolonged", "constant", "long-term"]
  },
  "chemistry": {
    "element": ["substance", "component", "constituent", "principle", "radix", "basic substance", "chemical element"],
    "compound": ["mixture", "amalgam", "synthesis", "combination", "concoction", "blend", "chemical compound"],
    "molecule": ["particle", "unit", "structure", "corpuscle", "molecular entity", "group of atoms"],
    "reaction": ["response", "process", "transformation", "interaction", "chemical change", "conversion", "synthesis"],
    "catalyst": ["agent", "stimulus", "impetus", "accelerant", "enzyme", "promoter", "reaction enhancer"],
    "solution": ["mixture", "blend", "suspension", "emulsion", "liquid", "solute-solvent mixture", "aqueous solution"],
    "bond": ["link", "connection", "attraction", "coupling", "linkage", "chemical link", "covalent bond"],
    "ion": ["charged particle", "anion", "cation", "charged atom", "electrolyte"],
    "acid": ["corrosive", "sour substance", "proton donor", "acidic compound"],
    "base": ["alkali", "proton acceptor", "antacid", "basic compound", "alkaline substance"],
    "alkane": ["paraffin", "saturated hydrocarbon", "methane series"],
    "alkene": ["olefin", "unsaturated hydrocarbon", "ethylene series"],
    "atom": ["particle", "corpuscle", "basic unit of matter", "elementary particle"],
    "concentration": ["strength", "density", "potency", "molarity", "amount of substance"],
    "distillation": ["purification", "refining", "extraction", "separation by boiling"],
    "electrode": ["anode", "cathode", "terminal", "conductor", "electrical conductor"],
    "electron": ["negatively charged particle", "lepton", "elementary particle"],
    "enzyme": ["biocatalyst", "protein catalyst", "ferment", "biological catalyst"],
    "equilibrium": ["balance", "stability", "state of balance", "chemical equilibrium"],
    "gas": ["vapor", "fume", "gaseous state", "volatile substance"],
    "hydrocarbon": ["organic compound", "alkane", "alkene", "compound of hydrogen and carbon"],
    "isotope": ["nuclide", "form of an element", "atomic variant"],
    "liquid": ["fluid", "solution", "liquefied substance", "aqueous state"],
    "metal": ["element", "conductor", "metallic substance", "alkali metal"],
    "neutron": ["subatomic particle", "nucleon", "neutral particle"],
    "nucleus": ["core", "center", "atomic nucleus", "central part"],
    "organic": ["carbon-based", "biological", "natural", "living"],
    "oxidation": ["corrosion", "rusting", "loss of electrons", "reaction with oxygen"],
    "pH": ["acidity", "alkalinity", "hydrogen ion concentration", "measure of acidity"],
    "polymer": ["macromolecule", "plastic", "resin", "large molecule chain"],
    "proton": ["positively charged particle", "nucleon", "hydrogen ion"],
    "reduction": ["gain of electrons", "deoxidation", "hydrogenation"],
    "salt": ["ionic compound", "saline substance", "product of acid-base reaction"],
    "solid": ["fixed shape substance", "crystalline solid", "non-fluid"],
    "solvent": ["dissolving liquid", "dissolvent", "liquid medium"],
    "substance": ["material", "matter", "compound", "element", "chemical"],
    "synthesis": ["creation", "formation", "production", "combination", "chemical creation"],
    "valence": ["combining capacity", "bonding capacity", "oxidation state"]
  },
  "literature": {
    "narrative": ["story", "account", "tale", "chronicle", "recital", "commentary", "plotline", "narration"],
    "character": ["persona", "figure", "protagonist", "individual", "type", "personage", "participant", "dramatis persona"],
    "theme": ["subject", "topic", "motif", "concept", "idea", "leitmotif", "message", "central idea"],
    "plot": ["storyline", "outline", "narrative", "scenario", "structure", "intrigue", "developments", "action"],
    "symbolism": ["representation", "imagery", "allegory", "metaphor", "connotation", "suggestion", "figurative meaning"],
    "author": ["writer", "creator", "novelist", "composer", "dramatist", "poet", "essayist", "wordsmith"],
    "metaphor": ["figure of speech", "analogy", "comparison", "symbol", "trope", "conceit", "implied comparison"],
    "prose": ["text", "composition", "writing", "narrative", "fiction", "discourse", "written language"],
    "irony": ["sarcasm", "paradox", "satire", "twist", "incongruity", "understatement", "double meaning"],
    "foreshadow": ["predict", "portend", "imply", "hint at", "allude to", "presage", "signal", "augur"],
    "allegory": ["parable", "fable", "symbolic story", "emblem", "extended metaphor"],
    "allusion": ["reference", "insinuation", "implication", "mention", "indirect reference"],
    "antagonist": ["adversary", "opponent", "villain", "nemesis", "rival", "foil"],
    "climax": ["peak", "pinnacle", "apex", "turning point", "culmination", "high point"],
    "conflict": ["struggle", "dispute", "clash", "tension", "opposition", "contention"],
    "dialogue": ["conversation", "discourse", "talk", "exchange", "repartee", "colloquy"],
    "genre": ["category", "class", "style", "type", "sort", "kind"],
    "imagery": ["description", "representation", "symbolism", "mental pictures", "sensory details"],
    "motif": ["theme", "idea", "pattern", "recurring element", "leitmotif", "dominant idea"],
    "personification": ["anthropomorphism", "humanization", "prosopopoeia", "attribution of human qualities"],
    "setting": ["backdrop", "environment", "location", "milieu", "surroundings", "context"],
    "stanza": ["verse", "canto", "strophe", "refrain", "verse paragraph"],
    "tone": ["mood", "atmosphere", "feeling", "attitude", "spirit", "inflection"],
    "tragedy": ["disaster", "calamity", "catastrophe", "downfall", "serious drama"],
    "comedy": ["humor", "satire", "farce", "light entertainment", "amusing play"],
    "epic": ["heroic poem", "long narrative", "saga", "grand tale"],
    "fable": ["parable", "moral tale", "allegory", "short story with a moral"],
    "folklore": ["mythology", "tradition", "lore", "legends", "traditional beliefs"],
    "legend": ["myth", "saga", "epic", "traditional story", "folk tale"],
    "myth": ["fable", "legend", "folktale", "parable", "traditional narrative"],
    "novel": ["book", "fiction", "story", "romance", "long-form narrative"],
    "poem": ["verse", "lyric", "ode", "sonnet", "rhyme", "composition in verse"],
    "sonnet": ["ballad", "lyric poem", "verse", "fourteen-line poem"],
    "verse": ["poetry", "rhyme", "stanza", "line", "metrical writing"],
    "act": ["division", "section", "part", "performance", "main division of a play"],
    "alliteration": ["head rhyme", "initial rhyme", "consonance", "repetition of initial sounds"],
    "anaphora": ["repetition", "epanaphora", "repetition of a word at the beginning of clauses"],
    "assonance": ["vowel rhyme", "vocalic rhyme", "repetition of vowel sounds"],
    "ballad": ["song", "poem", "chant", "narrative song", "simple narrative poem"],
    "connotation": ["implication", "suggestion", "undertone", "nuance", "associated meaning"],
    "denotation": ["literal meaning", "definition", "explicit meaning", "dictionary definition"],
    "diction": ["phrasing", "wording", "language", "terminology", "style", "choice of words"],
    "elegy": ["lament", "dirge", "requiem", "funeral song", "mournful poem"],
    "euphemism": ["understatement", "indirect term", "softening", "mild alternative"],
    "exposition": ["background", "introduction", "prelude", "setup", "explanatory part"],
    "hyperbole": ["exaggeration", "overstatement", "magnification", "extravagant statement"],
    "juxtaposition": ["comparison", "contrast", "collocation", "proximity", "placing side-by-side"],
    "mood": ["atmosphere", "ambiance", "feeling", "disposition", "emotional setting"],
    "onomatopoeia": ["imitative word", "sound symbolism", "echoism", "word that imitates a sound"],
    "oxymoron": ["contradiction", "paradox", "incongruity", "contradictory terms"],
    "parody": ["spoof", "satire", "caricature", "imitation", "comedic imitation"],
    "pathos": ["pity", "compassion", "sadness", "emotion", "quality that evokes pity"],
    "satire": ["irony", "mockery", "parody", "sarcasm", "wit used to critique"],
    "soliloquy": ["monologue", "discourse", "speech", "address", "act of speaking one's thoughts aloud"],
    "syntax": ["sentence structure", "arrangement", "grammar", "pattern", "word order"]
  },
  "business": {
    "strategy": ["plan", "approach", "policy", "tactic", "methodology", "blueprint", "game plan"],
    "revenue": ["income", "earnings", "turnover", "proceeds", "sales", "takings", "receipts"],
    "profit": ["gain", "return", "surplus", "yield", "bottom line", "net income", "earnings"],
    "market": ["clientele", "audience", "customer base", "industry", "sector", "demographic", "marketplace"],
    "brand": ["identity", "image", "logo", "reputation", "trademark", "name"],
    "stakeholder": ["shareholder", "investor", "partner", "participant", "contributor", "interested party"],
    "asset": ["property", "resource", "holding", "possession", "capital", "benefit"],
    "liability": ["debt", "obligation", "accountability", "burden", "disadvantage", "financial obligation"],
    "leverage": ["influence", "advantage", "power", "clout", "bargaining power", "strategic advantage"],
    "synergy": ["collaboration", "cooperation", "combined effort", "teamwork", "combined action"]
  }
};

const AI_TRACE_PATTERNS = [
  // Common Starters & Introductions
  "certainly, here('s| is)", "of course, here('s| is)", "sure, here('s| is)",
  "here('s| is) (a|an|the|your|what i found|what i came up with|how you can|how you could)",
  "here is a brief", "here is a summary", "here is an outline",
  "here is an introduction for you", "here's an introduction for you", 
  "here is a summary for you", "here's a summary for you",
  "certainly, i can help with that",
  "let’s break it down", "let me explain",
  "to begin with,?", "firstly,?", "first of all,?",

  // Summaries & Conclusions
  "in conclusion,", "in summary,", "to summarize,", "in sum,", "to sum up,", "overall,",
  "put simply,", "in short,", "all in all,", "the key takeaway is",
  "the main idea is", "ultimately,", "the bottom line is",

  // Polite Closings
  "(i )?hope this (helps|was helpful|is helpful|information helps|is useful)",
  "let me know if you have any (other|further) questions",
  "let me know if you need anything else",
  "feel free to (ask|reach out)",
  "please don’t hesitate to ask",

  // AI Identity / Limitations
  "as a large language model,", "as an ai language model,", "as an ai,",
  "as an artificial intelligence,", "i am an ai,", "i’m an ai,",
  "i (cannot|can't|am not able to|am unable to)",
  "i do not have the ability to", "i don’t have personal opinions",
  "i don’t have beliefs", "i do not have beliefs",
  "i don’t have personal experiences", "i lack personal experiences",
  "my knowledge cutoff is", "my training data only goes up to",
  "my knowledge is current up to",

  // Noting & Hedging Phrases
  "it is important to note", "it should be noted", "it’s worth noting that",
  "it is also important to note", "please note that",
  "however, it’s also important to consider",
  "it’s important to remember that",
  "keep in mind that", "one thing to keep in mind",
  "additionally,", "furthermore,", "moreover,",

  // Explaining / Teacher-Like Style
  "let’s go step by step", "let’s go through this",
  "to put it another way", "in other words,",
  "let’s break this down", "to clarify,",
  "this means that", "what this implies is",

  // Suggestion / Instruction Phrases
  "you could try", "you might consider",
  "one approach is", "another option is",
  "a common way to do this is", "a possible solution is",
  "an alternative is", "the recommended way is"
];
// ... [The rest of the Code.gs file remains the same]

