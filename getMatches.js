function getMatches() {
    const challongeApiKey =
        PropertiesService.getScriptProperties().getProperty("CHALLONGE_API_KEY");
    const tournamentId =
        PropertiesService.getScriptProperties().getProperty("TOURNAMENT_ID");
    var roundNum =
        PropertiesService.getScriptProperties().getProperty("ROUND_NUM");

    const options = {
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

    var responseTourney = getDetails(urlTourneyDetails, options);
    var responseRound = getDetails(urlRoundDetails, options);

    // getTourneyDetails(responseTourney);
    //   getMatchDetails(responseRound);
    round = completeRound(responseTourney, responseRound, roundNum);
    Logger.log(round);

    editSheet(round, roundNum);

    roundNum++;
    var updatedRoundNum = { "ROUND_NUM": roundNum };
    PropertiesService.getScriptProperties().setProperties(updatedRoundNum);

    editSheet(round, roundNum);

    updateMatchForm(round, roundNum);

    sendMail(roundNum);
}

function sendMail(roundNum) {
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

    var col = 'C';
    var emails = sheet.getRange(`${col}:${col}`).getValues();
    emails = emails.flat().filter(String)
    emails.shift();
    emails = emails.join(',');

    Logger.log(emails);
    var receipient = emails;


    var challongeLink = PropertiesService.getScriptProperties().getProperty('CHALLONGE_LINK');
    var sheetLink = PropertiesService.getScriptProperties().getProperty('INFO_LINK');
    var formLink = PropertiesService.getScriptProperties().getProperty('SUBMISSION_LINK');

    if (roundNum > 1) {
        var subject = `⚔️ ${roundNum} is Here — Your Next Challenge Awaits!`;
        var body = `
<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f4f4f4; padding: 20px;">
  <div style="background-color: #1a1a2e; padding: 30px; border-radius: 8px 8px 0 0; text-align: center;">
    <h1 style="color: #ffffff; margin: 0; font-size: 26px;">⚔️ ${roundNum}</h1>
    <p style="color: #e94560; margin: 8px 0 0; font-weight: bold;">The Battle Continues!</p>
  </div>

  <div style="background-color: #ffffff; padding: 30px; border-radius: 0 0 8px 8px;">
    <p style="font-size: 16px; color: #333;">You've made it to <strong>${roundNum}</strong> of the AUST CSE Carnival &lt;8.0/&gt; Chess Competition!</p>
    <p style="font-size: 16px; color: #333;">Ready to face your next challenger?</p>

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

    <div style="text-align: center; margin: 30px 0;">
      <a href="[Insert Google Form Link Here]" style="background-color: #e94560; color: #ffffff; padding: 14px 28px; border-radius: 6px; text-decoration: none; font-weight: bold; font-size: 16px; display: inline-block;">
        Report Your ${roundNum} Result
      </a>
    </div>

    <p style="font-size: 14px; color: #777; text-align: center;">Keep calm and calculate your moves — this round could decide everything!</p>
  </div>
</div>
`;
    }

    MailApp.sendEmail({ to: receipient, subject: subject, htmlBody: body });

}

function updateMatchForm(round, roundNum) {
    // Convert the round object into an array of match objects.
    const matches = Object.keys(round).map(key => {
        const [matchId, player1, player2, score] = round[key];
        return {
            matchNumber: letterToNumber(key),
            matchId: matchId != null ? String(matchId) : 'Unknown',
            player1: player1 || 'TBD',
            player2: player2 || 'TBD'
        };
    });

    // // Debug: log the exact data going into the form, useful if something looks off.
    // Logger.log(JSON.stringify(matches, null, 2));

    if (matches.length === 0) {
        Logger.log('No matches found for round ' + roundNum);
        return;
    }

    const formID = PropertiesService.getScriptProperties().getProperty('FORM_ID');
    let form;
    if (formID) {
        form = FormApp.openById(formID);

        // Step 1: Clear GO_TO_PAGE navigation on existing choices first.
        // Deleting page-break items while other items still reference them
        // via setGoToPage can throw "Invalid data updating form."
        const existingItems = form.getItems();
        existingItems.forEach(item => {
            if (item.getType() === FormApp.ItemType.LIST) {
                const listItem = item.asListItem();
                const plainChoices = listItem.getChoices().map(c => listItem.createChoice(c.getValue()));
                listItem.setChoices(plainChoices);
            } else if (item.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
                const mcItem = item.asMultipleChoiceItem();
                const plainChoices = mcItem.getChoices().map(c => mcItem.createChoice(c.getValue()));
                mcItem.setChoices(plainChoices);
            }
        });

        // Step 2: Now safe to delete all existing items, including page breaks.
        const items = form.getItems();
        for (let i = items.length - 1; i >= 0; i--) {
            form.deleteItem(items[i]);
        }
    } else {
        form = FormApp.create('Match Update');
        PropertiesService.getScriptProperties().setProperty('FORM_ID', form.getId());
    }
    form.setDescription('Select the match and report the winner.');

    const dropdownItem = form.addListItem();
    dropdownItem.setTitle('Which match was played?').setRequired(true);

    const matchLabels = matches.map(m => `${m.matchNumber} - ${m.matchId} (${m.player1} vs ${m.player2})`);
    const uniqueLabels = new Set(matchLabels);
    if (uniqueLabels.size !== matchLabels.length) {
        Logger.log('WARNING: duplicate match labels detected: ' + JSON.stringify(matchLabels));
    }

    const firstPageBreak = form.addPageBreakItem().setTitle('Match Details');

    const matchPageBreaks = [];
    const winnerItems = []; // hold refs so we can wire up navigation once finalPageBreak exists

    matches.forEach((match, i) => {
        let sectionStart;
        if (i === 0) {
            sectionStart = firstPageBreak; // first section right after the dropdown
        } else {
            sectionStart = form.addPageBreakItem().setTitle('Match ' + match.matchNumber);
        }
        matchPageBreaks.push(sectionStart);

        const winnerItem = form.addMultipleChoiceItem();
        winnerItem.setTitle(`Who won Match ${match.matchNumber}?`).setRequired(true);
        winnerItems.push({ item: winnerItem, match: match });
    });

    const finalPageBreak = form.addPageBreakItem()
        .setTitle('Thank you!')
        .setHelpText('Your response has been recorded.');

    const dropdownChoices = matchLabels.map((label, i) =>
        dropdownItem.createChoice(label, matchPageBreaks[i])
    );
    dropdownItem.setChoices(dropdownChoices);

    winnerItems.forEach(({ item, match }) => {
        item.setChoices([
            item.createChoice(match.player1, finalPageBreak),
            item.createChoice(match.player2, finalPageBreak),
            item.createChoice('Tie', finalPageBreak)
        ]);
    });

    Logger.log('Form updated: ' + form.getPublishedUrl());
}

function letterToNumber(letter) {
    return letter.toUpperCase().charCodeAt(0) - 64;
}

function editSheet(round, roundNum) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheetName = 'round' + roundNum;

    if (ss.getSheetByName(sheetName)) {
        ss.setActiveSheet(ss.getSheetByName(sheetName));
        ss.getActiveSheet().clearContents();
    }
    else {
        ss.insertSheet(sheetName);
        ss.setActiveSheet(ss.getSheetByName(sheetName));

    }

    const roundSheet = ss.getActiveSheet();
    Logger.log(roundSheet);

    // Set headers
    roundSheet.getRange(1, 1).setValue("Match No.")
    roundSheet.getRange(1, 2).setValue("Match ID");
    roundSheet.getRange(1, 3).setValue("Player 1");
    roundSheet.getRange(1, 4).setValue("Player 2");
    roundSheet.getRange(1, 5).setValue("Score");

    let row = 2;
    for (const matchId in round) {
        const matchDetails = round[matchId];
        matchNum = letterToNumber(matchId);
        roundSheet.getRange(row, 1).setValue(matchNum); // Match ID
        roundSheet.getRange(row, 2).setValue(matchDetails[0]); // Match ID
        roundSheet.getRange(row, 3).setValue(matchDetails[1]); // Player 1
        roundSheet.getRange(row, 4).setValue(matchDetails[2]); // Player 2
        roundSheet.getRange(row, 5).setValue(matchDetails[3]); // Score
        row++;
    }

}

function getDetails(url, options) {
    var result;
    try {
        const response = UrlFetchApp.fetch(url, options);
        const statusCode = response.getResponseCode();
        const content = response.getContentText();

        // Logger.log("HTTP Status: " + statusCode);
        // Logger.log("Response: " + content);

        if (statusCode >= 200 && statusCode < 300) {
            Logger.log("Received All Match Details");
            result = JSON.parse(content);
        } else {
            Logger.log("API Error");
            Logger.log(response);
        }
    } catch (error) {
        Logger.log("Fetch error");
        Logger.log(error);
    }
    return result;
}

function completeRound(tourney, match, roundNum) {
    var round = {};
    match.data.forEach((m) => {
        currMatch = m.attributes;
        if (currMatch.round == roundNum) {
            var matchDetails = {};
            var playerInfo = [];
            const matchAttrs = m.attributes;
            const p1Id = matchAttrs.points_by_participant[0].participant_id;
            Logger.log(currMatch.state);
            if (currMatch.state != 'complete') {
                putTie(m.id, p1Id);
            }
            var p1Name = getPlayers(tourney, p1Id);
            const p2Id = matchAttrs.points_by_participant[1].participant_id;
            var p2Name = getPlayers(tourney, p2Id);
            // Logger.log(
            //     `Match ID: ${m.id}, Player 1 ID: ${p1Id}, Name: ${p1Name}, Player 2 ID: ${p2Id}, Name: ${p2Name}`,
            // );
            playerInfo.push(m.id, p1Name, p2Name, matchAttrs.scores);
            round[matchAttrs.identifier] = playerInfo;
        }
    });
    return round;
}

function getRoundDetails(tourney, match, roundNum) {
    var round = {};
    match.data.forEach((m) => {
        currMatch = m.attributes;
        if (currMatch.round == roundNum) {
            var matchDetails = {};
            var playerInfo = [];
            const matchAttrs = m.attributes;
            const p1Id = matchAttrs.points_by_participant[0].participant_id;
            // Logger.log(currMatch.state);
            // if (currMatch.state != 'complete') {
            //     putTie(m.id, p1Id);
            // }
            var p1Name = getPlayers(tourney, p1Id);
            const p2Id = matchAttrs.points_by_participant[1].participant_id;
            var p2Name = getPlayers(tourney, p2Id);
            // Logger.log(
            //     `Match ID: ${m.id}, Player 1 ID: ${p1Id}, Name: ${p1Name}, Player 2 ID: ${p2Id}, Name: ${p2Name}`,
            // );
            playerInfo.push(m.id, p1Name, p2Name, matchAttrs.scores);
            round[matchAttrs.identifier] = playerInfo;
        }
    });
    return round;
}

function getPlayers(tourney, pID) {
    var foundName = null;
    tourney.data.forEach((participant) => {
        const attrs = participant.attributes;

        if (participant.id == pID) {
            foundName = attrs.name;
            return foundName;
        }
    });
    return foundName;
}

function putTie(matchId, participantId) {
    const challongeApiKey =
        PropertiesService.getScriptProperties().getProperty("CHALLONGE_API_KEY");
    const tournamentId =
        PropertiesService.getScriptProperties().getProperty("TOURNAMENT_ID");
    var urlMakeTie = `https://api.challonge.com/v2.1/tournaments/${tournamentId}/matches/${matchId}.json`;

    const payload = {
        data: {
            type: "match",
            attributes: {
                match: [
                    {
                        participant_id: participantId,
                        score_set: "0",
                        rank: 1,
                        advancing: false
                    }],
                tie: true
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
        const response = UrlFetchApp.fetch(urlMakeTie, options);
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
}