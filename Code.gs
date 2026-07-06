/**
 * RADAR Team Scoreboard - Backend Configuration & Functions
 * Conference: RADAR
 * Year: 2026
 */

// CONFIGURATION: Set your secure access code and sheet name here
const ACCESS_CODE = "RADAR@2026"; 
const SHEET_NAME = "Scores";

/**
 * Serves the web application HTML page with responsive adjustments.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('RADAR Team Scoreboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Helper function to retrieve or initialize the Google Sheet database.
 */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  // If the sheet doesn't exist, create it with default columns and sample data
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Team Name", "Score"]);
    sheet.appendRow(["Team Alpha", 0]);
    sheet.appendRow(["Team Beta", 0]);
    sheet.appendRow(["Team Gamma", 0]);
    sheet.appendRow(["Team Delta", 0]);
    
    // Format headers
    sheet.getRange("A1:B1").setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Fetches all unique team names dynamically from the Google Sheet.
 * @return {Array<string>} List of team names.
 */
function getTeams() {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const teams = [];
    
    // Start from index 1 to skip the header row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() !== "") {
        teams.push(data[i][0].toString().trim());
      }
    }
    return teams;
  } catch (e) {
    throw new Error("Failed to load teams from database: " + e.message);
  }
}

/**
 * Retrieves the team records, calculates rankings, and sorts descending.
 * @return {Array<Object>} Sorted list of team objects with ranking details.
 */
function getLeaderboard() {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const leaderboard = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() !== "") {
        leaderboard.push({
          name: data[i][0].toString().trim(),
          score: Number(data[i][1]) || 0
        });
      }
    }
    
    // Sort teams from highest score to lowest
    leaderboard.sort((a, b) => b.score - a.score);
    return leaderboard;
  } catch (e) {
    throw new Error("Failed to update leaderboard: " + e.message);
  }
}

/**
 * Updates a specific team's score after passing validation tests.
 * @param {string} teamName - Name of the team to update.
 * @param {number} points - Integer points to add/subtract (-50 to +50).
 * @param {string} accessCode - The code provided by the client.
 * @return {Object} Status object showing success or failure flags.
 */
function updateScore(teamName, points, accessCode) {
  // 1. Strict Backend Security Validation
  if (!accessCode || accessCode !== ACCESS_CODE) {
    return { success: false, message: "Access Denied: Invalid security access code!" };
  }
  
  // 2. Input Sanitization & Numeric Parsing
  const parsedPoints = parseInt(points, 10);
  if (isNaN(parsedPoints)) {
    return { success: false, message: "Validation Error: Points must be a valid integer." };
  }
  
  // 3. Score Boundary Enforcement
  if (parsedPoints < -50 || parsedPoints > 50) {
    return { success: false, message: "Validation Error: Points value must be between -50 and +50." };
  }
  
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    let rowUpdated = -1;
    let currentScore = 0;
    
    // Find matching team row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === teamName.trim()) {
        rowUpdated = i + 1; // +1 to convert zero-based loop index to sheet row index
        currentScore = Number(data[i][1]) || 0;
        break;
      }
    }
    
    if (rowUpdated === -1) {
      return { success: false, message: "Error: Selected team was not found in database." };
    }
    
    // Write new cumulative value to the sheet
    const newScore = currentScore + parsedPoints;
    sheet.getRange(rowUpdated, 2).setValue(newScore);
    
    return { 
      success: true, 
      message: `Score updated! ${teamName} received ${parsedPoints > 0 ? '+' : ''}${parsedPoints} points.` 
    };
    
  } catch (e) {
    return { success: false, message: "Database Transaction Error: " + e.message };
  }
}
