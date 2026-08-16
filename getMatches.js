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
    round = getRoundDetails(responseTourney, responseRound, roundNum);
    Logger.log(round);

    editSheet(round, roundNum);
    roundNum++;
    var updatedRoundNum = { "ROUND_NUM": roundNum };
    PropertiesService.getScriptProperties().setProperties(updatedRoundNum);
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

function getRoundDetails(tourney, match, roundNum) {
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
                        participant_id: "302057535",
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