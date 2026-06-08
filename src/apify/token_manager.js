/**
 * 1. ONE-TIME SETUP
 * Insert your three tokens here, run this function once from the script editor.
 * After successful execution, you can remove the tokens from the code for security.
 * If ever any problem with tokens, run this again.

function initializeTokensOneTime() {
  var props = PropertiesService.getScriptProperties();
  
  // INSERT YOUR TOKENS HERE:
  props.setProperty('APIFY_TOKEN_1', 'insert token 1');
  props.setProperty('APIFY_TOKEN_2', 'insert token 2');
  props.setProperty('APIFY_TOKEN_3', 'insert token  3');
  
  props.setProperty('ACTIVE_TOKEN_INDEX', '1');
  props.setProperty('RUN_COUNT_1', '0');
  props.setProperty('RUN_COUNT_2', '0');
  props.setProperty('RUN_COUNT_3', '0');
  
  console.log("Tokens successfully saved. Current active token: 1");
}
*/

/**
 * 2. BUTTON FUNCTION FOR THE SHEET
 * It switches the token in a 1 -> 2 -> 3 -> 1 cycle and resets the counter to 0.
 */
function switchApifyToken() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var currentIndex = props.getProperty('ACTIVE_TOKEN_INDEX') || '1';
  
  // Determine the next token in the sequence
  var newIndex;
  if (currentIndex === '1') {
    newIndex = '2';
  } else if (currentIndex === '2') {
    newIndex = '3';
  } else {
    newIndex = '1';
  }
  
  // Reset the counter for the token we are switching to
  props.setProperty('RUN_COUNT_' + newIndex, '0');
  props.setProperty('ACTIVE_TOKEN_INDEX', newIndex);
  
  ui.alert(
    "Token Switch", 
    "Successfully switched to Token " + newIndex + ".\nRun counter for Token " + newIndex + " reset to 0.", 
    ui.ButtonSet.OK
  );
}

/**
 * 3. GET ACTIVE TOKEN
 * Use this function instead of CONFIG.APIFY.TOKEN
 */
function getActiveApifyToken() {
  var props = PropertiesService.getScriptProperties();
  var activeIndex = props.getProperty('ACTIVE_TOKEN_INDEX') || '1';
  var token = props.getProperty('APIFY_TOKEN_' + activeIndex);
  
  if (!token) {
    throw new Error("Token not found. Please run initializeTokensOneTime().");
  }
  return token;
}

/**
 * 4. LOG THE RUN
 * Call this function at the very end of your data import script.
 * It increments the counter by 1, checks limits, and displays a message.
 */
function logRunAndCheckLimits() {
  var props = PropertiesService.getScriptProperties();
  var activeIndex = props.getProperty('ACTIVE_TOKEN_INDEX') || '1';
  var currentRuns = parseInt(props.getProperty('RUN_COUNT_' + activeIndex) || '0', 10);
  
  // Increment the counter
  currentRuns += 1;
  props.setProperty('RUN_COUNT_' + activeIndex, currentRuns.toString());
  
  // Limit math
  var maxRuns = 17; // 5.00 / 0.28 = 17.85
  var ui = SpreadsheetApp.getUi();
  var msg = "Runs on current token (" + activeIndex + "): " + currentRuns + " out of " + maxRuns;
  
  if (currentRuns >= 18) {
    ui.alert("CRITICAL", "Credit limit reached! " + msg + ".\nPlease switch the token.", ui.ButtonSet.OK);
  } else if (currentRuns >= 15) {
    ui.alert("WARNING", "Credits are running low. " + msg + ".\nA token switch will be needed soon.", ui.ButtonSet.OK);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Import complete");
  }
}
