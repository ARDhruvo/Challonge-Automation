function getMatches() {
  const challongeApiKey =
    PropertiesService.getScriptProperties().getProperty("CHALLONGE_API_KEY");
  const tournamentId =
    PropertiesService.getScriptProperties().getProperty("TOURNAMENT_ID");

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
  Logger.log(getRoundDetails(responseTourney, responseRound));
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

function getRoundDetails(tourney, match) {
  var round = {};
  match.data.forEach((m) => {
    var matchDetails = {};
    var playerInfo = [];
    const matchAttrs = m.attributes;
    const p1Id = matchAttrs.points_by_participant[0].participant_id;
    var p1Name = getPlayers(tourney, p1Id);
    const p2Id = matchAttrs.points_by_participant[1].participant_id;
    var p2Name = getPlayers(tourney, p2Id);
    Logger.log(
      `Match ID: ${m.id}, Player 1 ID: ${p1Id}, Name: ${p1Name}, Player 2 ID: ${p2Id}, Name: ${p2Name}`,
    );
    playerInfo.push(m.id, p1Name, p2Name, matchAttrs.scores);
    round[matchAttrs.identifier] = playerInfo;
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

function getTourneyDetails(data) {
  const invitations = {};
  data.included.forEach((inv) => {
    invitations[inv.id] = inv.attributes;
  });

  data.data.forEach((participant) => {
    const attrs = participant.attributes;
    Logger.log(`Seed ${attrs.seed}: ${attrs.name} (ID: ${participant.id})`);
    Logger.log(`  Active: ${attrs.states.active}`);
    Logger.log(`  Username: ${attrs.username}`);

    const invData = participant.relationships.invitation.data;
    if (invData) {
      const inv = invitations[invData.id];
      Logger.log(`  Invitation ID: ${invData.id}`);
      Logger.log(`  Invitation accepted: ${inv.accepted}`);
    }

    Logger.log("---");
  });
}

function getMatchDetails(data) {
  data.data.forEach((match) => {
    const attrs = match.attributes;

    Logger.log(`Match ID: ${match.id}`);
    Logger.log(`Identifier: ${attrs.identifier}`);
    Logger.log(`Round: ${attrs.round}`);
    Logger.log(`State: ${attrs.state}`);
    Logger.log(`Score: ${attrs.scores}`);

    attrs.points_by_participant.forEach((p) => {
      Logger.log(`  Participant ID: ${p.participant_id}`);
      Logger.log(`  Scores: ${p.scores}`);
    });
  });
}
