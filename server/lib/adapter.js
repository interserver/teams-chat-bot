// Shared BotFrameworkAdapter singleton.
//
// All proactive-send paths (msgController, dailyRecapController,
// notificationConsumer, botActivityHandler sync) share one adapter so
// credentials and the token cache are pooled in one place.

const { BotFrameworkAdapter } = require('botbuilder');

let _adapter = null;

function buildAdapter() {
    return new BotFrameworkAdapter({
        appId: process.env.MicrosoftAppId,
        appPassword: process.env.MicrosoftAppPassword
    });
}

function getAdapter() {
    if (!_adapter) {
        _adapter = buildAdapter();
        _adapter.onTurnError = async (context, error) => {
            console.error(`[shared adapter onTurnError] ${ error }`);
        };
    }
    return _adapter;
}

// Allow tests/consumer to replace the adapter (e.g. with a mock)
function setAdapter(adapter) {
    _adapter = adapter;
}

module.exports = { getAdapter, setAdapter };
