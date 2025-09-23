/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("processBtn").addEventListener("click", processDocument);
  }
});

async function processDocument() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      
      const settings = {
        watermark: document.getElementById('chkWatermark').checked,
        spaceInsertion: document.getElementById('chkSpaceInsertion').checked,
        spacePunctuation: document.getElementById('spacePunctuation').value,
        spaceInsertionProb: parseFloat(document.getElementById('spaceInsertionProb').value)
      };

      let modifiedText = body.text;
      
      if (settings.watermark) {
        modifiedText = modifiedText.replace(/[\u200B-\u200D\uFEFF]/g, '').replace(/\s{2,}/g, ' ');
      }
      
      if (settings.spaceInsertion) {
        const punc = settings.spacePunctuation.split('');
        let tempText = "";
        for(let i = 0; i < modifiedText.length; i++) {
          let char = modifiedText[i];
          if (i > 0 && punc.includes(char) && modifiedText[i-1] !== ' ' && Math.random() < settings.spaceInsertionProb) {
            tempText += ' ';
          }
          tempText += char;
        }
        modifiedText = tempText;
      }

      body.insertText(modifiedText, Word.InsertLocation.replace);
      await context.sync();
      
      document.getElementById('status').textContent = "Processing complete!";
    });
  } catch (error) {
    console.error(error);
    document.getElementById('status').textContent = "Error: " + error.message;
  }
}

