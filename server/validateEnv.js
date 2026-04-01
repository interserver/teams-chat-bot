// Validates required environment variables at startup.
// Throws if any required variable is missing.

const REQUIRED = [
    'MicrosoftAppId',
    'MicrosoftAppPassword',
    'MYSQL_HOST',
    'MYSQL_USER',
    'MYSQL_PASS',
    'MYSQL_DB',
    'ZONEMTA_USERNAME',
    'ZONEMTA_PASSWORD',
    'ZONEMTA_HOST'
];

function validateEnv() {
    const missing = REQUIRED.filter(key => !process.env[key]);
    if (missing.length > 0) {
        throw new Error(`Missing required environment variables: ${ missing.join(', ') }`);
    }
}

module.exports = { validateEnv, REQUIRED };
