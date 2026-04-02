const { MessageFactory } = require('botbuilder');

module.exports = {
    match(text, lcText, { ima }) {
        if (ima !== 'admin') return null;
        return /^hyperv status$/i.test(lcText) ? {} : null;
    },
    async execute(_match, { context, db, redis }) {
        try {
            const [rows] = await db.query(
                `SELECT * FROM vps_masters
                LEFT JOIN vps_master_details USING (vps_id)
                WHERE vps_type = (SELECT type_id FROM vps_platform_types WHERE type_key = 'HYPERV' LIMIT 1)`
            );
            if (!rows || rows.length === 0) {
                await context.sendActivity(MessageFactory.text('No HyperV hosts found.'));
                return;
            }
            const now = Math.floor(Date.now() / 1000);
            const inUse = [];
            for (const row of rows) {
                const hostKey = `vps_host_${ row.vps_id }`;
                const requestKey = `${ hostKey }_request`;
                const ts = await redis.get(hostKey);
                if (ts && ts !== '0') {
                    const request = await redis.get(requestKey) || 'unknown';
                    inUse.push(`${ row.vps_name }#${ row.vps_id } for ${ now - parseInt(ts, 10) }s on ${ request }`);
                }
            }
            const msg = `${ inUse.length } hosts ${ inUse.length > 0
                ? `(${ inUse.join(', ') })`
                : '' } in the middle of processing a queue`;
            await context.sendActivity(MessageFactory.text(msg));
        } catch (err) {
            await context.sendActivity(MessageFactory.text(`Error checking HyperV status: ${ err.message }`));
            console.error('HyperV status error:', err);
        }
    }
};
