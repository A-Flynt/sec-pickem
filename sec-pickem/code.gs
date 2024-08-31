/*
Author: Austin Flynt

This script autofills the SEC pick em. Gets data using the college football data API
Link: https://api.collegefootballdata.com/api/docs/?url=/api-docs.json#/games/
*/


/*

Function:     getSeasonData

Returns:      None

Description:  This function fills the data sheet with the time, teams, and score if available.
              It will also populate a list of SEC teams, organized by division, along with their records and rank if applicable

*/
function getSeasonData(player_sheet, data_sheet, season, conference) 
{

  // Set the data sheet as the active sheet and clear old data
  var sheet = SpreadsheetApp.getActive().getSheetByName(data_sheet);
  sheet.activate();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet.hideSheet();
  sheet.clearContents()

  // Build the appropriate url and fetch the data with the API call
  url = `https://api.collegefootballdata.com/games?year=${season}&seasonType=regular&conference=${conference}`
  var response = UrlFetchApp.fetch(url, {
  headers: {accept: "application/json",Authorization: "Bearer J0lxRoeUXu+U2eftZ7fwVVcGuC330glGAk81vGKrwSJk5XJzkvSLPpLgRpcFLZ6U"}
  });

  // Parse the text into JSON
  var json = response.getContentText();
  var mae = JSON.parse(json); 

  // Set week flags to determine what week it is
  week = 0;
  this_week = 0;
  week_set = false;

  // Create range for sorting
  top_left = 1;
  bottom_right = 1;

  //Create range for sub-sorting
  sub_top_left = top_left;

  // Create flags for sub-sort (same day and time games) 
  same_time = false;
  timestamp = -1;

  // Get todays's timestamp
  today = Date.parse(Date());
  odds = "-";

  // For each game a team in the requested conference plays this week
  for (var i = 0; i < mae.length; i++) {

    // Start empty array for row and get the info for game i
    var stats=[]; 
    var game = mae[i];

    // If we've moved into the next week, make a break in the sheet, sort last week
    if(week != game.week)
    {
      // Moves range up a row to only capture scores and sort range by kickoff time
      bottom_right = bottom_right - 1;
      if(week == 0)
      {
        top_left = 3
        bottom_right = 3
      }
      else
      {
        range_str = `A${top_left}:F${bottom_right}`;
        range = sheet.getRange(range_str);
        range.sort([6,3]);
        top_left = bottom_right+3;
        bottom_right = top_left;
      }


      // Make break in the sheet
      week = game.week;
      ss.appendRow(['Total']);
      str_week = `Week ${game.week}`;
      ss.appendRow([str_week,"Away", "Home", "Odds", "Score"]);
    }

    // Format game time to make more human readable
    game_time = new Date(game.start_date)
    last_timestamp = timestamp;
    timestamp = Date.parse(game_time);

    // Get the spread of the game
    odds = getOdds(game.id, data_sheet, season);

    same_time = last_timestamp == timestamp;
    const formattedDate = game_time.toLocaleString("en-US", {
      day: "numeric",
      month: "short",
      year: "numeric",
      hour: "numeric",
      minute: "2-digit"
    });

    // Build row with needed info
    stats.push(formattedDate);
    stats.push(game.away_team);
    stats.push(game.home_team);
    stats.push(odds)
    home_p = game.home_points;
    away_p = game.away_points;

    // If home team won, show them as winner and append score
    if(home_p > away_p) 
    {
      score_str = `${game.home_team} (${home_p}-${away_p})`;
      stats.push(score_str);
    } 

    // If away team won, show them as winner and append score
    else if(away_p > home_p) 
    {
      score_str = `${game.away_team} (${away_p}-${home_p})`;
      stats.push(score_str);
    } 

    // If tie, say tie and append score
    else if(away_p == home_p && home_p != undefined)
    {
      score_str = `Tie (${home_p}-${away_p})`;
      stats.push(score_str);
    } 

    // Otherwise put a '-' for no score 
    else 
    {
      stats.push("-")
      if(!week_set)
      {
        this_week = week - 1;
        week_set = true;
      }
    }

    this_week = week;

    // Append row built with this games info
    stats.push(timestamp)
    ss.appendRow(stats);

    // Extends the range to sort
    bottom_right = bottom_right + 1;
  } // End for loop


  schools = ["Alabama", "Ole Miss","Arkansas","Mississippi State","Texas A&M","Auburn","LSU", "Georgia", "Kentucky", "Tennessee", "Missouri", "South Carolina", "Florida", "Vanderbilt", "Texas", "Oklahoma"]; 

  // Build conference table - Full conferance column
  var cell = ss.getRange("N13"); 
  cell.setValue("Rankings");

  for(var i = 0; i < schools.length; i++)
  {
    // Get the stats for team i and move to next row
    row = 14+i;
    var cell = ss.getRange(`N${row}`); 
    team_stats = getStats(schools[i],this_week, data_sheet, season, conference) ;
    cell.setValue(team_stats);
  } // End for loop

  // TODO: Remove old division based code

  // Arrays for SEC divisions 
  //west = ["Alabama", "Ole Miss","Arkansas","Mississippi State","Texas A&M","Auburn","LSU"];
  //east = ["Georgia", "Kentucky", "Tennessee", "Missouri", "South Carolina", "Florida", "Vanderbilt"];

  //// Build conference table - West column
  //var cell = ss.getRange("N13"); 
  //cell.setValue("West");

  // For each team in division
  //for(var i = 0; i < west.length; i++)
  //{
  //  // Get the stats for team i and move to next row
  //  row = 14+i;
  //  var cell = ss.getRange(`N${row}`); 
  //  team_stats = getStats(west[i],this_week, data_sheet, season, conference) ;
  //  cell.setValue(team_stats);
  //} // End for loop

  //// Build conference table - East column
  //var cell = ss.getRange("O13"); 
  //cell.setValue("East");

  // For each team in division
  //for(var i = 0; i < east.length; i++)
  //{
  //  // Get the stats for team i and move to next row
  //  row = 14+i;
  //  var cell = ss.getRange(`O${row}`); 
  //  team_stats = getStats(east[i],this_week, data_sheet, season, conference) ;
  //  cell.setValue(team_stats);
  //} // End for loop

} // End getSeasonData



/*

Function:     getRanking

Returns:      Integer

Description:  This function returns the rank of a team
              
*/
function getRanking(team_arg, week_arg, data_sheet, season, conference)
{
  // Remove this when rankings are released
  // return undefined

  // Set data sheet as active sheet and hide the sheet
  var sheet = SpreadsheetApp.getActive().getSheetByName(data_sheet);
  sheet.activate();
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //get active spreadsheet (bound to this script)
  sheet.hideSheet();

  // Build the url and fetch the ranking data using the constructed API call
  var url = `https://api.collegefootballdata.com/rankings?year=${season}&week=${week_arg}&seasonType=regular`
  var response = UrlFetchApp.fetch(url, {
  headers: {accept: "application/json",Authorization: "Bearer J0lxRoeUXu+U2eftZ7fwVVcGuC330glGAk81vGKrwSJk5XJzkvSLPpLgRpcFLZ6U"}
  });

  // Parse the return text into JSON
  var json = response.getContentText();
  var mae = JSON.parse(json);
  Logger.log(url)
  // Retrive AP Poll Rankings
  ranks_req = mae[0].polls[0].ranks

  // For each rank
  for(var i = 0; i < ranks_req.length; i++)
  {
    // If the school matches school argument, get ranking
    if(ranks_req[i].school == team_arg)
    {
      return ranks_req[i].rank
    }
  }
} // End getRanking



/*

Function:     getStats

Returns:      String

Description:  This function retrieves the most up-to-date rank and record for a team
              
*/
function getStats(team_arg, week_arg, data_sheet, season, conference)
{
  // Set data sheet as active sheet and hide the sheet
  var sheet = SpreadsheetApp.getActive().getSheetByName(data_sheet);
  sheet.activate();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet.hideSheet();

  // Encode to handle any special characters like ampersands
  uri_encoded_team = encodeURIComponent(team_arg);

  // Build the url and fetch the records data using the constructed API call
  url = `https://api.collegefootballdata.com/records?year=${season}&team=${uri_encoded_team}&conference=${conference}`
  var response = UrlFetchApp.fetch(url, {
  headers: {accept: "application/json",Authorization: "Bearer J0lxRoeUXu+U2eftZ7fwVVcGuC330glGAk81vGKrwSJk5XJzkvSLPpLgRpcFLZ6U"}
  });

  // Parse the return text into JSON
  var json = response.getContentText(); 
  var mae = JSON.parse(json); 

  // Get team stats and find ranking if applicable
  var team_stats = mae[0];
  rank = getRanking(team_arg, week_arg, data_sheet, season, conference);

  // If the team isn't ranked, leave rank blank
  if(rank == undefined)
  {
    rank = ''
  }
  return `${rank} ${team_arg} (${team_stats.total.wins}-${team_stats.total.losses})`;
    
}


/*

Function:     getOdds

Returns:      String

Description:  This function retrieves the odds of each game
              
*/
function getOdds(id, data_sheet, season)
{
  // Set data sheet as active sheet and hide the sheet
  var sheet = SpreadsheetApp.getActive().getSheetByName(data_sheet);
  sheet.activate();
  var ss = SpreadsheetApp.getActiveSpreadsheet(); //get active spreadsheet (bound to this script)
  sheet.hideSheet();

  // Build the url and fetch the ranking data using the constructed API call
  var url = `https://api.collegefootballdata.com/lines?gameId=${id}&year=${season}`
  var response = UrlFetchApp.fetch(url, {
  headers: {accept: "application/json",Authorization: "Bearer J0lxRoeUXu+U2eftZ7fwVVcGuC330glGAk81vGKrwSJk5XJzkvSLPpLgRpcFLZ6U"}
  });

  // Parse the return text into JSON
  var json = response.getContentText();
  var mae = JSON.parse(json);

  // Retrive Spreads
  for(var k = 0; k < mae.length; k++)
  {
    if(mae[k].id == id)
    {
      game = mae[k].lines[0];

      if(game != undefined)
      {
        return mae[k].lines[0].formattedSpread
      }
      else
      {
        return "-"
      }
    }
  }
} // End getOdds


/*

Function:     countColoredCells

Returns:      int

Description:  This function returns the number of cells of a certain color
              From this site: https://spreadsheetpoint.com/count-cells-based-on-cell-color-google-sheets/
              
*/
function countColoredCells(countRange,colorRef) {
  var activeRange = SpreadsheetApp.getActiveRange();
  var activeSheet = activeRange.getSheet();
  var formula = activeRange.getFormula();
  
  var rangeA1Notation = formula.match(/\((.*)\,/).pop();
  var range = activeSheet.getRange(rangeA1Notation);
  var bg = range.getBackgrounds();
  var values = range.getValues();
  
  var colorCellA1Notation = formula.match(/\,(.*)\)/).pop();
  var colorCell = activeSheet.getRange(colorCellA1Notation);
  var color = colorCell.getBackground();
  
  var count = 0;
  
  for(var i=0;i<bg.length;i++)
    for(var j=0;j<bg[0].length;j++)
      if( bg[i][j] == color )
        count=count+1;
  return count;
}; // End countColoredCells


/*

Function:     main

Returns:      None

Description:  This function gets settings from the settings sheet and gets the season data
              based on these settings
              
*/
function main() {

  // Set settings sheet as active sheet
  var sheet = SpreadsheetApp.getActive().getSheetByName('Settings');
  sheet.activate();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get settings values
  var player_sheet = ss.getRange("A2").getValue();
  var data_sheet = ss.getRange("B2").getValue();
  var season = ss.getRange("C2").getValue();
  var conference = ss.getRange("D2").getValue(); 
  // TODO: USE DIVISION INFO TO GET GENERATE RANKS. CURRENTLY HARD CODED. SHOULD WAIT UNTIL NEW TEAMS ADDED TO CONFERENCE

  // Populate data sheet
  getSeasonData(player_sheet, data_sheet, season, conference);

}
