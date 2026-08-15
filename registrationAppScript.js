function onFormSubmit(e) {
    const formResponse = e.response;
    const itemResponses = formResponse.getItemResponses();

    // Set your Challonge API key and tournament ID in the script properties inside Project Settings > Script Properties
    const challongeApiKey = PropertiesService.getScriptProperties().getProperty('CHALLONGE_API_KEY');
    const tournamentId = PropertiesService.getScriptProperties().getProperty('TOURNAMENT_ID');

    const playerName = itemResponses[0].getResponse();  // Name
    // const playerEmail = itemResponses[1].getResponse(); // Email

    const payload = {
        data: {
            type: "participant",
            attributes: {
                name: playerName,
                seed: 1,
                misc: "",
                // email: playerEmail
            }
        }
    };

    const options = {
        method: 'POST',
        headers: {
            'Accept': 'application/json',
            'Authorization-Type': 'v1',
            'Authorization': challongeApiKey,
            'Content-Type': 'application/vnd.api+json'
        },
        payload: JSON.stringify(payload),   // <-- IT HAS TO BE PAYLOAD!!!!!!!!!!!!!
        muteHttpExceptions: true,
        redirect: 'follow'
    };
    const url = `https://api.challonge.com/v2.1/tournaments/${tournamentId}/participants.json`;

    try {
        const response = UrlFetchApp.fetch(url, options);
        const statusCode = response.getResponseCode();
        const content = response.getContentText();

        Logger.log('HTTP Status: ' + statusCode);
        Logger.log('Response: ' + content);

        if (statusCode >= 200 && statusCode < 300) {
            Logger.log('Participant added successfully: ' + playerName + ' (' + playerEmail + ')');
        } else {
            Logger.log('API error for ' + playerName + ' (' + playerEmail + '): ' + content);
        }
    } catch (error) {
        Logger.log('Fetch error for ' + playerName + ' (' + playerEmail + '): ' + error.toString());
    }
}