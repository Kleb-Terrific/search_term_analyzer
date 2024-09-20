// Main function to evaluate search terms
function evaluateSearchTerms(spreadsheet, startRow, endRow) {
  //If a function calls evaluateSearchTerms function without passing the Spreadsheet parameter
  if (!spreadsheet) {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    if (!spreadsheet) {
      Logger.log('Error: Spreadsheet not found');
      return;
    }
  }
  
  const analysisSearchTermSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.ANALYSIS_SEARCH_TERM);
  const promptPageSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.PROMPT_PAGE);
  const planningPageSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.PLANNING_PAGE);

  if (!(analysisSearchTermSheet && promptPageSheet && planningPageSheet)) {
    Logger.log(`Error: Check Sheet Names' spelling and spacing`); return;
  }

  Logger.log(`Analyzing: Rows  ${startRow} until ${endRow}`);
  const searchTermColumn = 'A';
  const searchTermColumnNum = 0;
  const decisionColumnNum = 6; //Column G
  
  const searchTerms = measureExecutionTime('getRangeValues', () => getRangeValues(analysisSearchTermSheet, startRow, endRow, searchTermColumn));
  const promptParts = measureExecutionTime('getPromptParts', () => getPromptParts(promptPageSheet));
  const negKeywords = measureExecutionTime('getNegativeKeywords', () => getNegativeKeywords(planningPageSheet));

  measureExecutionTime('filterNegativeKeywords', () => 
    filterNegativeKeywords(analysisSearchTermSheet, searchTerms, startRow, decisionColumnNum, negKeywords));

  const searchTermsFiltered = measureExecutionTime('getFilteredKeywords', () => 
    getFilteredKeywords(analysisSearchTermSheet, startRow, endRow, searchTermColumnNum, decisionColumnNum));

  const evaluations = measureExecutionTime('getEvaluationFromChatGPT', () => getEvaluationFromChatGPT(searchTermsFiltered, promptParts));

  
  if (evaluations && evaluations.result) {
    measureExecutionTime('updateSheetWithEvaluations', () => 
      updateSheetWithEvaluations(analysisSearchTermSheet, searchTermsFiltered, evaluations.result, decisionColumnNum));
  }

  //Analysis - Search Terms Cost and Impressions Formula Setting
  setCostAndImpressionFormula(analysisSearchTermSheet);
}


function filterNegativeKeywords(analysisSearchTermSheet, searchTerms, startRow, decisionColumnNum, negKeywords){
  searchTerms.forEach((row, index) => {
    const keyword = row.trim(); 
    let matchType = ""; 
    let culpritKeyword  = "";

    // Check for Negative Exact Match
    if (negKeywords["Exact Match"].some(negKeyword => {
      if (negKeyword.toLowerCase() === keyword.toLowerCase()) {
        culpritKeyword = negKeyword;
        return true;
      }
      return false;
    })) {
      matchType = "Negative Exact Match: ";
    }  

    // Check for Negative Phrase Match
    else if (negKeywords["Phrase Match"].some(negKeyword => {
      const negKeywordLower = negKeyword.toLowerCase();
      const keywordLower = keyword.toLowerCase();
      if (keywordLower.includes(negKeywordLower) && isPhraseMatch(keywordLower, negKeywordLower)) {
        culpritKeyword = negKeyword;
        return true;
      }
      return false;
    })) {
      matchType = "Negative Phrase Match: ";
    } 

    // Check for Negative Broad Match
    else if (negKeywords["Broad Match"].some(negKeyword => {
      const negWords = negKeyword.toLowerCase().split(' ');
      const keywordWords = keyword.toLowerCase().split(' ');
      if (negWords.every(word => keywordWords.includes(word))) {
        culpritKeyword = negKeyword;
        return true;
      }
      return false;
    })) {
      matchType = "Negative Broad Match: ";
    }

    // Create JSON output
    const jsonOutput = {
      search_term: keyword,
      decision: matchType ? "EXCLUDE" : "",
      confidence: matchType ? "100%" : "",
      reason: matchType + culpritKeyword
    };

    analysisSearchTermSheet.getRange(startRow + index, decisionColumnNum+1).setValue(jsonOutput.decision);
    analysisSearchTermSheet.getRange(startRow + index, decisionColumnNum+2).setValue(jsonOutput.confidence);
    analysisSearchTermSheet.getRange(startRow + index, decisionColumnNum+3).setValue(jsonOutput.reason);
  });
}


function getFilteredKeywords(analysisSearchTermSheet, startRow, endRow, searchTermColumnNum, decisionColumnNum){
  const values = analysisSearchTermSheet.getRange(startRow, 1, endRow - startRow + 1, 7).getValues();

  let filteredKeywords = [];
  let offset = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (row[decisionColumnNum] !== "EXCLUDE") {
      offset = 0;
      filteredKeywords.push({
        rowNumber: i + startRow + offset, 
        value: row[searchTermColumnNum]
    });
    ;}else offset = offset + 1; 
  }
  return filteredKeywords;
}


// Helper function to check if the keyword contains the phrase in the correct order
function isPhraseMatch(keyword, negKeyword) {
  const keywordWords = keyword.split(' ');
  const negKeywordWords = negKeyword.split(' ');

  let keywordIndex = 0;

  for (let i = 0; i < negKeywordWords.length; i++) {
    keywordIndex = keywordWords.indexOf(negKeywordWords[i], keywordIndex);
    if (keywordIndex === -1) {
      return false; // If the negKeyword word isn't found in the correct order, return false
    }
    keywordIndex++; // Move to the next word in the keyword
  }
  
  return true; // All negKeyword words are found in the correct order
}

// Helper function to get values from a specified range
function getRangeValues(sheet, startRow, endRow, column) {
  return sheet.getRange(`${column}${startRow}:${column}${endRow}`).getValues().flat();
}

// Fetch prompt parts from the promptPageSheet
function getPromptParts(promptPageSheet){
  try{
    return{
      companyName: promptPageSheet.getRange('A2').getValue(),
      companyWebsite: promptPageSheet.getRange('B2').getValue(),	
      companyOverview: promptPageSheet.getRange('C2').getValue(),
      companyProducts: promptPageSheet.getRange('D2').getValue(),
      competitors: promptPageSheet.getRange('E2').getValue(),
      rules: promptPageSheet.getRange('F2').getValue()
    }
    
  } catch(error){
    Logger.log(`Error: ${error.message}`);
    return{ companyName: [], companyWebsite: [], companyOverview: [], companyProducts: [], competitors: [], rules: []}
  }
}

// Fetch negative keywords from the provided Google Sheets URL
function getNegativeKeywords(planningPageSheet) {
  try {    
    const dataRange = planningPageSheet.getRange(2, 1, planningPageSheet.getLastRow() - 1, 2); // Assumes header is in row 1
    const data = dataRange.getValues();

    const negKeywordDict = {
      "Broad Match": [],
      "Phrase Match": [],
      "Exact Match": []
    };

    data.forEach(row => {
    const negKeyword = row[0];
    const matchType = row[1];
    
    switch (matchType.toLowerCase().trim()) {
      case "broad":
        negKeywordDict["Broad Match"].push(negKeyword.trim());
        break;
      case "exact":
        negKeywordDict["Exact Match"].push(negKeyword.trim());
        break;
      case "phrase":
        negKeywordDict["Phrase Match"].push(negKeyword.trim());
        break;
    }
  });

    return negKeywordDict;
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    return negKeywordDict = {"Broad Match": [], "Phrase Match": [],"Exact Match": []};
  }
}

// Helper function to get values from a specified column - Assuming Row 1 is a header
function getColumnValues(sheet, column) {
  return sheet.getRange(`${column}2:${column}${sheet.getLastRow()}`).getValues().flat().filter(Boolean);
}

// Function to interact with the ChatGPT API for evaluation
function getEvaluationFromChatGPT(searchTerms, promptParts) {
  const prompt = generatePrompt(searchTerms, promptParts);
  //Logger.log(prompt);

  try {
    const response = UrlFetchApp.fetch(API_URL, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${OPENAI_API_KEY}`,
        'Content-Type': 'application/json',
      },
      payload: JSON.stringify({
        model: MODEL,
        messages: [
          { role: "system", content: "You are an expert Google Ads and Search Campaign Manager that evaluates search terms based on its relevance to the company's goods, services, niche, and/or offerings" },
          { role: "user", content: prompt }
        ],
        temperature: 0.0, //0.7 - Rup's suggestion = slower
      }),
    });

    const gptResult = JSON.parse(response.getContentText());
    Logger.log("Tokens Used: " + gptResult.usage.total_tokens)
    return JSON.parse(gptResult.choices[0].message.content.trim());
  } catch (error) {
    Logger.log(`Error calling OpenAI API: ${error.message}`);
    return null;
  }
}

// Function to generate the prompt for the ChatGPT API
function generatePrompt(searchTerms, promptParts) {
  return `
    //Persona
    You are an expert Google Ads Search Campaign Manager for the company ${promptParts.companyName}.
    You speak fluent English, German, Spanish, Chinese, Japanese, Portuguese, and Hebrew.
    You are meticulous, consistent, and accurate in your analysis, providing concise yet effective evaluations.
    You never hallucinate. You do not make up factual information. You focus on relevant information.
    You create campaigns using the company's brand and its competitors' keywords.

    //Company Context & Overview
    ${promptParts.companyOverview}

    //Task
    Evaluate if a search term should be included or excluded in the Google Ad campaign based on company's products, services, and niche industry.

    //Special Rules
    ${promptParts.rules}

    //General Rules
    1. Include terms that directly relate to company's products, services, and niche.
    2. Include terms RELATED to CLOSE COMPETITORS and of: ${promptParts.competitors} if they are aligned with company's products, services, and niche.
    3. Exclude generic terms unless paired with relevant keywords.
    4. Exclude unrelated terms.

    //Output
      1. Search Term
      1. Decision = "INCLUDE" or "EXCLUDE"
      2. Confidence Percent = Confidence percentage of your decision
      3. Reason = A one-sentence summary of your decision

    ##STRICT OUTPUT FORMAT - DO NOT WRAP THE JSON CODES IN JSON MARKERS
    {
      "result": [
        {"search_term": "", "decision": "", "confidence": "", "reason": "" },
        {"search_term": "", "decision": "", "confidence": "", "reason": "" }
      ]
    }

    Search Terms: ${searchTerms.map(item => item.value).join(', ')}`;
}

// Helper function to update the sheet with evaluation results
function updateSheetWithEvaluations(analysisSearchTermSheet, searchTermsFiltered, evaluations, decisionColumnNum) {

  const searchTermsDictionary = searchTermsFiltered.reduce((acc, item) => {
    acc[item.value] = Math.floor(item.rowNumber); // Using Math.floor to convert float to integer
    return acc;
  }, {});

  evaluations.forEach((item) => {
    const searchTerm = item.search_term;
    const rowNumber = searchTermsDictionary[searchTerm];

    if (rowNumber) {
      // If a matching row number is found, populate the cells.
      analysisSearchTermSheet.getRange(rowNumber, decisionColumnNum+1).setValue(item.decision);
      analysisSearchTermSheet.getRange(rowNumber, decisionColumnNum+2).setValue(item.confidence);
      analysisSearchTermSheet.getRange(rowNumber, decisionColumnNum+3).setValue(item.reason);
    } else { Logger.log(`No matching row found for search term: ${searchTerm}`);}
  });
}

function setCostAndImpressionFormula(analysisSearchTermSheet){
  spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  analysisSearchTermSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.ANALYSIS_SEARCH_TERM);
  lastRow = analysisSearchTermSheet.getLastRow();
  const formulaCost = '=SUMIF(\'All Search Terms\'!A:A,A2,\'All Search Terms\'!B:B)';
  const formulaImpression = '=SUMIF(\'All Search Terms\'!A:A,A2,\'All Search Terms\'!C:C)';

  analysisSearchTermSheet.getRange(`J2:J${lastRow}`).setFormula(formulaCost);
  analysisSearchTermSheet.getRange(`K2:K${lastRow}`).setFormula(formulaImpression);
}


//Google Sheets UI
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Search Term Analyzer')
    .addItem('Keyword Analyzer', 'showDialog')
    .addItem('Keyword Analyzer (until last word)', 'showDialogv2')
    .addItem('Run N-Gram Analyzer', 'runNGramAnalyzer')
    .addItem('Clear Analysis - Search Terms Sheet', 'clearAnalysisSearchTerms')
    .addToUi();
}


function showDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('dialog')
    .setWidth(400)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Keyword Analyzer');
}

function showDialogv2() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('dialogv2')
    .setWidth(400)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Keyword Analyzer (until last word)');
}

function clearAnalysisSearchTerms() {
  var ui = SpreadsheetApp.getUi();
  
  // Display a confirmation dialog
  var response = ui.alert('Confirm Clear',
                          'Are you sure you want to clear the content in A2:L of "Analysis - Search Terms" sheet? This action cannot be undone.',
                          ui.ButtonSet.YES_NO);

  // Check the user's response
  if (response == ui.Button.YES) {
    var spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.ANALYSIS_SEARCH_TERM);
    
    if (sheet) {
      sheet.getRange('A2:L').clearContent();
      ui.alert('Content cleared from A2:L in "Analysis - Search Terms" sheet.');
    } else {
      ui.alert('Sheet "Analysis - Search Terms" not found.');
    }
  } else {
    ui.alert('Operation cancelled. No data was cleared.');
  }
}


function runEvaluation(startRow = null, endRow = null) {
  let spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const defaultStartRow = 2;
  const defaultEndRow = CONFIG.BATCH_SIZE; // Batches of 15

  // Use default values if parameters are not provided
  const effectiveStartRow = startRow !== null ? startRow : defaultStartRow;
  const effectiveEndRow = endRow !== null ? endRow : (startRow !== null ? effectiveStartRow + CONFIG.BATCH_SIZE : defaultEndRow);

    // Create a simple HTML message
  const htmlOutput = HtmlService.createHtmlOutput('<p>Keyword Analyzer (Specific Rows)</p>')
    .setWidth(250)
    .setHeight(80);

  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, 'Keyword Analyzer: Rows ' + effectiveStartRow + ' until ' + effectiveEndRow);

  if (effectiveEndRow - effectiveStartRow > CONFIG.BATCH_SIZE)
    processRowsInBatches(effectiveStartRow, effectiveEndRow)
  else
    evaluateSearchTerms(spreadsheet, effectiveStartRow, effectiveEndRow);
}

function runEvaluationNoDialog(startRow = null, endRow = null) {
  let spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const defaultStartRow = 2;
  const defaultEndRow = CONFIG.BATCH_SIZE; // Batches of 15

  // Use default values if parameters are not provided
  const effectiveStartRow = startRow !== null ? startRow : defaultStartRow;
  const effectiveEndRow = endRow !== null ? endRow : (startRow !== null ? effectiveStartRow + CONFIG.BATCH_SIZE : defaultEndRow);

  if (effectiveEndRow - effectiveStartRow > CONFIG.BATCH_SIZE)
    processRowsInBatchesNoDialog(effectiveStartRow, effectiveEndRow)
  else
    evaluateSearchTerms(spreadsheet, effectiveStartRow, effectiveEndRow);
}


//Automated Batch Processing until last row (by 15 rows)
function processRowsInBatches(startRow, givenlastRow) {  
  let spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
  let analysisSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.ANALYSIS_SEARCH_TERM);

  if (!analysisSheet) {
    SpreadsheetApp.getUi().alert('Sheet not found!');
    return;
  }
  
  if (startRow == null){
    startRow = 2;
  }
  const lastRow = givenlastRow == null ? analysisSheet.getLastRow() : givenlastRow; // Get the last row with data in the sheet
  
  // Initialize the current row to start processing
  let currentRow = startRow;
  
  // Loop through the rows in batches of 15
  while (currentRow <= lastRow) {
    const endRow = Math.min(currentRow + CONFIG.BATCH_SIZE - 1, lastRow); // Calculate the end row for the current batch
    
    const htmlOutput = HtmlService.createHtmlOutput('<p>Keyword Analyzer (until last row)</p>')
    .setWidth(250)
    .setHeight(80);

    // Call your existing function to process rows within the current batch
    SpreadsheetApp.getUi().showModelessDialog(htmlOutput, `Processing Rows: ${currentRow} to ${endRow}`);
    evaluateSearchTerms(spreadsheet, currentRow, endRow);
    
    // Move to the next batch
    currentRow += CONFIG.BATCH_SIZE;
  }
  SpreadsheetApp.getUi().alert('Running N-Gram Analyzer');
  runNGramAnalyzer();
  SpreadsheetApp.getUi().alert('Batch processing completed.');
  
}

function processRowsInBatchesNoDialog(startRow, givenlastRow) {  
  let spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
  let analysisSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.ANALYSIS_SEARCH_TERM);

  if (!analysisSheet) {
    Logger.log("Sheet not found");
    return;
  }
  
  if (startRow == null){
    startRow = 2;
  }
  const lastRow = givenlastRow == null ? analysisSheet.getLastRow() : givenlastRow; // Get the last row with data in the sheet
  
  // Initialize the current row to start processing
  let currentRow = startRow;
  
  // Loop through the rows in batches of 15
  while (currentRow <= lastRow) {
    const endRow = Math.min(currentRow + CONFIG.BATCH_SIZE - 1, lastRow); // Calculate the end row for the current batch
    
    evaluateSearchTerms(spreadsheet, currentRow, endRow);
    
    // Move to the next batch
    currentRow += CONFIG.BATCH_SIZE;
  }
  setCostAndImpressionFormula();
  runNGramAnalyzer(); 
}


//For Unigram Analysis
function runNGramAnalyzer() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const analysisNGramSheet = spreadsheet.getSheetByName('Analysis - N Gram');
  const searchTermsSheet = spreadsheet.getSheetByName('Analysis - Search Terms');

  if (!analysisNGramSheet || !searchTermsSheet) {
    Logger.log('Error: Required sheets not found');
    return;
  }

  // Remove existing filter if any
  let existingFilter = analysisNGramSheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }

  // Clear existing content
  analysisNGramSheet.clear();

  // Set headers
  analysisNGramSheet.getRange('A1:F1').setValues([['N-Gram', 'Exclude Count', 'Total Unique Count', 'Exclude Ratio','Cost (last 7 days)', 'Impression (last 7 days)']]);

  // Get all search terms
  const searchTerms = searchTermsSheet.getRange('A2:A' + searchTermsSheet.getLastRow()).getValues().flat();
  
  // Get unique N-grams
  const nGrams = getUniqueNGrams(searchTerms);

  // Set N-grams in column A
  analysisNGramSheet.getRange(2, 1, nGrams.length, 1).setValues(nGrams.map(ng => [ng]));

  // Set formulas for columns B, C, and D
  const lastRow = analysisNGramSheet.getLastRow();
  const formulaB = '=COUNTIF(FILTER(\'Analysis - Search Terms\'!G:G, REGEXMATCH(\'Analysis - Search Terms\'!A:A,"\\b\\s?" & A2 & "\\s?\\b")),"EXCLUDE")';
  const formulaC = '=COUNTIF(FILTER(\'Analysis - Search Terms\'!A:A, REGEXMATCH(\'Analysis - Search Terms\'!A:A,"\\b\\s?" & A2 & "\\s?\\b")), "<>")';
  const formulaD = '=IF(B2<>"", B2 / C2, "")';
  const formulaE = '=SUM(FILTER(\'All Search Terms\'!B:B, REGEXMATCH(\'All Search Terms\'!A:A,"\\b\\s?" & A2 & "\\s?\\b")))';
  const formulaF = '=SUM(FILTER(\'All Search Terms\'!C:C, REGEXMATCH(\'All Search Terms\'!A:A,"\\b\\s?" & A2 & "\\s?\\b")))';

  analysisNGramSheet.getRange(`B2:B${lastRow}`).setFormula(formulaB);
  analysisNGramSheet.getRange(`C2:C${lastRow}`).setFormula(formulaC);
  analysisNGramSheet.getRange(`D2:D${lastRow}`).setFormula(formulaD);
  analysisNGramSheet.getRange(`E2:E${lastRow}`).setFormula(formulaE);
  analysisNGramSheet.getRange(`F2:F${lastRow}`).setFormula(formulaF);

  // Format as a table
  analysisNGramSheet.getRange(1, 1, lastRow, 6).createFilter();
  analysisNGramSheet.getRange('D2:D').setNumberFormat('0.00%');
}

function getUniqueNGrams(searchTerms) {
  const allWords = searchTerms.join(' ').toLowerCase().split(/\s+/);
  const uniqueWords = [...new Set(allWords)];
  return uniqueWords.filter(word => word.trim() !== '');
}

function measureExecutionTime(functionName, callback) {
  const start = Date.now();
  const result = callback();
  const end = Date.now();
  const executionTime = end - start;
  Logger.log(`${functionName} execution time: ${executionTime} ms`);
  return result;
}

function autoRun(){
  processRowsInBatchesNoDialog(2)
}