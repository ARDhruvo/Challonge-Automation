function getMatches() {
    const challongeApiKey = PropertiesService.getScriptProperties().getProperty('CHALLONGE_API_KEY');
    const tournamentId = PropertiesService.getScriptProperties().getProperty('TOURNAMENT_ID');

    const options = {
        method: 'GET',
        headers: {
            'Accept': 'application/json',
            'Authorization-Type': 'v1',
            'Authorization': challongeApiKey,
            'Content-Type': 'application/vnd.api+json'
        },
        muteHttpExceptions: true,
        redirect: 'follow'
    };
    const url = `https://api.challonge.com/v2.1/tournaments/${tournamentId}/matches.json`;

    try {
        const response = UrlFetchApp.fetch(url, options);
        const statusCode = response.getResponseCode();
        const content = response.getContentText();

        Logger.log('HTTP Status: ' + statusCode);
        Logger.log('Response: ' + content);

        if (statusCode >= 200 && statusCode < 300) {
            Logger.log('Received All Match Details');
        } else {
            Logger.log('API Error');
            Logger.log(response);
        }
    } catch (error) {
        Logger.log('Fetch error');
    }
}
