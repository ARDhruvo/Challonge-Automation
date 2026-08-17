function updateResult(e) {
    const formResponse = e.response;
    const itemResponses = formResponse.getItemResponses();
    const email = e.response.getRespondentEmail();

    if (!canSubmitToday(email)) {
        Logger.log(`Blocked duplicate submission from ${email} for today.`);
        var subject = "Submission Not Accepted";
        var body = `<h1>Your response was not accepted</h1>
        <p>You already submitted a response today. Please contact organizers if you made a mistake</p>
        <p>Thank you!</p>`

        MailApp.sendEmail({ to: email, subject: subject, htmlBody: body });
        return; // stop processing entirely
    }
    recordSubmission(email);

    // Logger.log(itemResponses);
    Logger.log(itemResponses[0].getResponse());
    var winner = itemResponses[1].getResponse();
    Logger.log(winner);
    // Logger.log(email);

    var players = extractNames(itemResponses[0].getResponse());
    // Logger.log(players[0]);


    const challongeApiKey =
        PropertiesService.getScriptProperties().getProperty("CHALLONGE_API_KEY");
    const tournamentId =
        PropertiesService.getScriptProperties().getProperty("TOURNAMENT_ID");

    if (winner == 'Tie' || winner == players[0]) {
        var matchDeets = getMatch(players[0]);
        matchId = matchDeets[0];
        participantId = matchDeets[1];
    }
    else {
        var matchDeets = getMatch(players[1]);
        matchId = matchDeets[0];
        participantId = matchDeets[1];
    }


    // Logger.log(matchId);

    var urlUpdate = `https://api.challonge.com/v2.1/tournaments/${tournamentId}/matches/${matchId}.json`;

    var score;
    var isTie;
    if (winner == 'Tie') {
        score = "0"
        isTie = true;
    }
    else {
        score = "1"
        isTie = false;

    }

    const payload = {
        data: {
            type: "match",
            attributes: {
                match: [
                    {
                        participant_id: participantId,
                        score_set: score,
                        rank: 1,
                        advancing: !isTie
                    }],
                tie: isTie
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
        const response = UrlFetchApp.fetch(urlUpdate, options);
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

function getMatch(player) {

    const challongeApiKey =
        PropertiesService.getScriptProperties().getProperty("CHALLONGE_API_KEY");
    const tournamentId =
        PropertiesService.getScriptProperties().getProperty("TOURNAMENT_ID");
    const roundNum =
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


    var tourney = getDetails(urlTourneyDetails, options)
    var responseRound = getDetails(urlRoundDetails, options);

    // Logger.log(responseRound);


    var matchDeets = [];

    responseRound.data.forEach((m) => {
        currMatch = m.attributes;
        if (currMatch.round == roundNum) {
            const matchAttrs = m.attributes;
            const p1Id = matchAttrs.points_by_participant[0].participant_id;
            var p1Name = getPlayers(tourney, p1Id);
            if (p1Name == player) {
                matchDeets[0] = m.id;
                matchDeets[1] = p1Id;
                return matchDeets;
            }
            const p2Id = matchAttrs.points_by_participant[1].participant_id;
            var p2Name = getPlayers(tourney, p2Id);
            if (p2Name == player) {
                matchDeets[0] = m.id;
                matchDeets[1] = p2Id;
                return matchDeets;
            }
        }
    });
    return matchDeets;

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

function extractNames(input) {
    // Match: anything before the first '(', then capture everything until the last ')'
    const regex = /^[^\(]+\((.+)\)$/;
    const match = input.match(regex);

    if (match) {
        const inside = match[1]; // "asdasd (invitation pending) vs asdsd (invitation pending)"
        return inside.split(' vs ');
    }
    return null;
}

function getTodayKey_() {
    // e.g. "2026-08-17" in your script's timezone
    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function canSubmitToday(email) {
    const props = PropertiesService.getScriptProperties();
    const key = "submitted_" + email + "_" + getTodayKey_();
    return props.getProperty(key) === null;
}

function recordSubmission(email) {
    const props = PropertiesService.getScriptProperties();
    const key = "submitted_" + email + "_" + getTodayKey_();
    props.setProperty(key, "1");
}