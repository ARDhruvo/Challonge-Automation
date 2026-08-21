function startTournament() {
    const ss = SpreadsheetApp.getActiveSpreadsheet()

    const sheetName = 'Form Responses 1';

    if (ss.getSheetByName(sheetName)) {
        ss.setActiveSheet(ss.getSheetByName(sheetName));
    }
    else {
        ss.insertSheet(sheetName);
        ss.setActiveSheet(ss.getSheetByName(sheetName));

    }

    var sheet = ss.getActiveSheet();

    var col = 'B';
    var emails = sheet.getRange(`${col}:${col}`).getValues();
    emails = emails.flat().filter(String)
    emails.shift();
    emails = emails.join(',');

    var challongeLink = PropertiesService.getScriptProperties().getProperty('CHALLONGE_LINK');
    var sheetLink = PropertiesService.getScriptProperties().getProperty('INFO_LINK');
    var formLink = PropertiesService.getScriptProperties().getProperty('SUBMISSION_LINK');

    Logger.log(emails);
    var receipient = emails;

    var subject = `♟️ Checkmate Season Begins — AUST CSE Carnival <8.0/> Chess is LIVE!`;
    var body = `
<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f4f4f4; padding: 20px;">
  <div style="background-color: #1a1a2e; padding: 30px; border-radius: 8px 8px 0 0; text-align: center;">
    <h1 style="color: #ffffff; margin: 0; font-size: 26px;">♟️ AUST Chess Has Begun!</h1>
  </div>

  <div style="background-color: #ffffff; padding: 30px; border-radius: 0 0 8px 8px;">
    <p style="font-size: 16px; color: #333;">The wait is over — the board is set, and it's time to move your pieces!</p>
    <p style="font-size: 16px; color: #333;">The <strong>AUST CSE Carnival &lt;8.0/&gt; Chess Competition</strong> has officially started. Here's everything you need:</p>

    <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
      <tr>
        <td style="padding: 12px; border-bottom: 1px solid #eee;">🏆 <strong>Tournament Bracket</strong></td>
        <td style="padding: 12px; border-bottom: 1px solid #eee; text-align: right;">
          <a href="${challongeLink}" style="color: #1a1a2e; font-weight: bold;">View Bracket →</a>
        </td>
      </tr>
      <tr>
        <td style="padding: 12px; border-bottom: 1px solid #eee;">📋 <strong>Player Info Sheet</strong></td>
        <td style="padding: 12px; border-bottom: 1px solid #eee; text-align: right;">
          <a href="${sheetLink}" style="color: #1a1a2e; font-weight: bold;">Check Details →</a>
        </td>
      </tr>
      <tr>
        <td style="padding: 12px;">📝 <strong>Match Update Form</strong></td>
        <td style="padding: 12px; text-align: right;">
          <a href="${formLink}" style="color: #1a1a2e; font-weight: bold;">Open Form →</a>
        </td>
      </tr>
    </table>

    <p style="font-size: 16px; color: #333;">Check the bracket for your first opponent, get in touch with them, and play your match!</p>

    <div style="text-align: center; margin: 30px 0;">
      <a href="[Insert Google Form Link Here]" style="background-color: #e94560; color: #ffffff; padding: 14px 28px; border-radius: 6px; text-decoration: none; font-weight: bold; font-size: 16px; display: inline-block;">
        Submit Your Result
      </a>
    </div>

    <p style="font-size: 14px; color: #777; text-align: center;">Good luck — may the best mind win!</p>
  </div>
</div>
`;

    MailApp.sendEmail({ to: receipient, subject: subject, htmlBody: body });


    // ---- 1. Create the destination spreadsheet (or reuse if it already exists) ----
    const destSheetName = 'Player Info';
    const destFileName = 'Player Info <8.0/> Chess';
    const props = PropertiesService.getScriptProperties();

    let destFileId = props.getProperty('SHEET_ID');
    let destSpreadsheet;

    if (destFileId) {
        try {
            destSpreadsheet = SpreadsheetApp.openById(destFileId);
        } catch (e) {
            destFileId = null; // stored id is stale — fall through and create a fresh one
        }
    }

    if (!destFileId) {
        destSpreadsheet = SpreadsheetApp.create(destFileName);
        props.setProperty('SHEET_ID', destSpreadsheet.getId());
        Logger.log(`Created new spreadsheet: ${destSpreadsheet.getUrl()}`);
    }

    let destSheet = destSpreadsheet.getSheetByName(destSheetName);
    if (!destSheet) {
        destSheet = destSpreadsheet.insertSheet(destSheetName);
    } else {
        destSheet.clearContents();
    }

    // ---- 2. Copy columns B:C and E:H from source into destination ----
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) {
        Logger.log('Source sheet is empty — nothing to copy.');
        return;
    }

    const bc = sheet.getRange(1, 2, lastRow, 2).getValues(); // B:C
    const eh = sheet.getRange(1, 5, lastRow, 4).getValues(); // E:H
    const combined = bc.map((row, i) => row.concat(eh[i]));

    destSheet.getRange(1, 1, combined.length, combined[0].length).setValues(combined);
    destSheet.autoResizeColumns(1, combined[0].length);

    Logger.log(`Copied ${combined.length} rows to "${destFileName}" > "${destSheetName}".`);
    Logger.log(`Player Info Sheet URL: ${destSpreadsheet.getUrl()}`);


    var challongeApiKey = props.getProperty('CHALLONGE_API_KEY');
    var tournamentId = props.getProperty('TOURNAMENT_ID');
    // var tournamentId = props.getScriptProperties('DUMMY_TOURNAMENT');
    var urlStart = `https://api.challonge.com/v2.1/tournaments/${tournamentId}/change_state.json`

    // const raw = JSON.stringify({
    //     "data": {
    //         "type": "TournamentState",
    //         "attributes": {
    //             "state": "process_checkin"
    //         }
    //     }
    // });

    const payload = {
        data: {
            type: "TournamentState",
            attributes: {
                state: "start"
            }
        }
    };

    const options = {
        method: "PUT",
        headers: {
            Accept: "application/json",
            "Authorization-Type": "v1",
            Authorization: challongeApiKey,
            "Content-Type": "application/vnd.api+json",
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
        redirect: "follow",
    };

    try {
        const response = UrlFetchApp.fetch(urlStart, options);
        const statusCode = response.getResponseCode();
        const content = response.getContentText();

        Logger.log("HTTP Status: " + statusCode);
        Logger.log("Response: " + content);

        if (statusCode >= 200 && statusCode < 300) {
            Logger.log("Updated as a tie");
            result = JSON.parse(content);
        } else {
            Logger.log("API Error");
            Logger.log(response);
        }
    } catch (error) {
        Logger.log("Fetch error");
        Logger.log(error);
    }

    // ---- 3. Set up Round 1 ----
    // This used to live in getMatches.js's roundNum==1 branch. It's moved
    // here so getMatches.js (running on its own daily trigger from
    // TOURNAMENT_START_DATE + 1 onward) only ever has to handle "complete
    // the previous round, set up the next one" — no separate first-round
    // case. getDetails/getRoundDetails/editSheet/updateMatchForm are
    // defined in getMatches.js, which lives in this same script project,
    // so they're callable directly here.
    // Note: this needs its own GET options — the `options` above is the
    // PUT request used for the Challonge state-change call.
    const getOptions = {
        method: "GET",
        headers: {
            Accept: "application/json",
            "Authorization-Type": "v1",
            Authorization: challongeApiKey,
            "Content-Type": "application/vnd.api+json",
        },
        muteHttpExceptions: true,
        redirect: "follow",
    };
    const urlTourneyDetails = `https://api.challonge.com/v2.1/tournaments/${tournamentId}/participants.json`;
    const urlRoundDetails = `https://api.challonge.com/v2.1/tournaments/${tournamentId}/matches.json`;

    var round1Tourney = getDetails(urlTourneyDetails, getOptions);
    var round1Matches = getDetails(urlRoundDetails, getOptions);

    var round1 = getRoundDetails(round1Tourney, round1Matches, 1);
    editSheet(round1, 1);
    updateMatchForm(round1, 1);

    PropertiesService.getScriptProperties().setProperty('ROUND_NUM', '2');
    Logger.log('Round 1 set up. ROUND_NUM initialized to 2 for getMatches.js.');
}