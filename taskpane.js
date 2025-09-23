/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// --- BUILT-IN DICTIONARY (Copied from your Code.gs) ---
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
const AI_TRACE_PHRASES = [
  // Common Starters & Introductions
  "certainly, here is", "certainly, here's", "of course, here is", "of course, here's",
  "here is a", "here's a", "here is an", "here's an", "here is your", "here's your",
  "here is the", "here's the", "sure, here is", "sure, here's",
  "here is a brief", "here is a summary", "here is an outline",
  "here is an introduction for you", "here's an introduction for you", 
  "here is a summary for you", "here's a summary for you",
  "certainly, i can help with that",

  // Common Endings & Closings
  "in conclusion,", "in summary,", "to summarize,", "in sum,", "to sum up,", "overall,",
  "i hope this helps", "i hope this is helpful", "i hope this was helpful",
  "i hope this information is helpful",
  "hope this helps", "hope this is helpful",
  "let me know if you have any other questions", "let me know if you need anything else",
  "feel free to ask", "feel free to reach out", "if you have any other questions",
  "if you have any more questions", "if you need further assistance", 
  "please let me know if you have any other questions",

  // AI Disclaimers & Identity Phrases
  "as a large language model,", "as an ai language model,", "as an ai,","I can suggest","I cannot suggest","I can't suggest","I suggest",
  "as an artificial intelligence,", "i am an ai,", "i'm an ai,", "i can't directly","pages", "i can generate", "i can help you",
  "i am not able", "i am unable", "i do not have the ability to", "I cannot provide","I can't provide","I can't", "i cannot", "i can",
  "i don't have personal opinions", "i don't have beliefs", "i do not have beliefs","real-time news","updates","browse",
  "i do not have personal experiences", "i lack personal experiences",
  "my knowledge cutoff is", "my knowledge is current up to",
  "my training data only goes up to",
  "it is important to note", "it should be noted", "it's worth noting that",
  "it is also important to note", "please note that",
  
  // Conversational Filler
  "however, it's also important to consider",
  "it's important to remember that",
  "additionally,"
];


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
                // The search API in Word can be case sensitive, so we handle both cases for the first letter
                let searchPhrase = phrase;
                if (/[a-zA-Z]/.test(phrase.charAt(0))) {
                    searchPhrase = `[${phrase.charAt(0).toUpperCase()}${phrase.charAt(0).toLowerCase()}]${phrase.slice(1)}`;
                }
                const searchResults = body.search(searchPhrase, { matchWildcards: true });
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
            const color = document.getElementById('highlightColor').value;
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load("items/font");
            await context.sync();
            let count = 0;

            for (const p of paragraphs.items) {
                // This is a simplified approach. A more robust solution would iterate through text runs within each paragraph.
                // For now, we search for highlighted text and delete it.
                // It's tricky to find all highlighted text directly, so we search for the phrases again and check color.
                for (const phrase of AI_TRACE_PHRASES) {
                    let searchPhrase = `[${phrase.charAt(0).toUpperCase()}${phrase.charAt(0).toLowerCase()}]${phrase.slice(1)}`;
                    const searchResults = p.search(searchPhrase, { matchWildcards: true });
                    context.load(searchResults, 'font');
                    await context.sync();
                    
                    for (const item of searchResults.items) {
                       if (item.font.highlightColor === color) {
                           item.delete();
                           count++;
                       }
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
                if (!originalText || originalText.trim() === "") continue;

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
  const homoglyphs = { 'a': 'а', 'e': 'е', 'o': 'о', 'c': 'с', 'i': 'і', 'p': 'р', 's': 'ѕ', 'x': 'х', 'A': 'А', 'E': 'Е', 'O': 'О', 'C': 'С', 'I': 'І', 'P': 'Р', 'S': 'Ѕ', 'X': 'Х' };
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
    const usToUk = {"analyze":"analyse","behavior":"behaviour","center":"centre","color":"colour","defense":"defence","favorite":"favourite","flavor":"flavour","gray":"grey","humor":"humour","labor":"labour","license":"licence","neighbor":"neighbour","organize":"organise","realize":"realise","recognize":"recognise","theater":"theatre","traveled":"travelled"};
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

