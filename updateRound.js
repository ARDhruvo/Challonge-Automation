function updateRound() {
    // Fires daily at midnight. Only increment ROUND_NUM while the
    // tournament is actually running, and stop on/after the final day —
    // there's no round after the last one, so nothing further to sync
    // with getMatches.js in the other project.
    if (!isTournamentActiveDay_()) {
        Logger.log('Outside tournament window (' + getTodayDateOnly_() + ') — ROUND_NUM not incremented.');
        return;
    }
    if (isLastTournamentDay_()) {
        Logger.log('Final tournament day — ROUND_NUM not incremented further.');
        return;
    }

    roundNum = PropertiesService.getScriptProperties().getProperty('ROUND_NUM');
    roundNum++;
    var update = { 'ROUND_NUM': String(roundNum) };
    PropertiesService.getScriptProperties().setProperties(update);
}

// ---- Date-gating helpers ----
// Requires two script properties on THIS project (Match Updates Form):
// TOURNAMENT_START_DATE and TOURNAMENT_END_DATE, both in "yyyy-MM-dd"
// format, e.g. "2026-09-06" and "2026-09-11". Separate Apps Script
// project from getMatches.js, so these need to be set here too, matching
// the values used there.

function getTodayDateOnly_() {
    var tz = Session.getScriptTimeZone();
    return Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
}

function isTournamentActiveDay_() {
    var props = PropertiesService.getScriptProperties();
    var start = props.getProperty('TOURNAMENT_START_DATE');
    var end = props.getProperty('TOURNAMENT_END_DATE');
    if (!start || !end) {
        Logger.log('TOURNAMENT_START_DATE / TOURNAMENT_END_DATE script properties are not set — refusing to run.');
        return false;
    }
    var today = getTodayDateOnly_();
    // Mirrors getMatches.js's window in the other project: ROUND_NUM
    // advances from the day after kickoff through the end date.
    return today > start && today <= end;
}

function isLastTournamentDay_() {
    var props = PropertiesService.getScriptProperties();
    var end = props.getProperty('TOURNAMENT_END_DATE');
    return getTodayDateOnly_() === end;
}