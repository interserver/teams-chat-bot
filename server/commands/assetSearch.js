const { MessageFactory } = require('botbuilder');

module.exports = {
    match(text, lcText, { ima }) {
        if (ima !== 'admin') return null;
        const m = text.match(/^(?:search|find|lookup|locate)\s+asset\s+(.+)$/i);
        return m ? { query: m[1].trim() } : null;
    },
    async execute({ query }, { context, db }) {
        const isNumeric = /^\d+$/.test(query);
        const sql = `SELECT *, assets.id AS real_asset_id
            FROM assets
            LEFT JOIN asset_types ON type_id = asset_types.asset_id
            LEFT JOIN asset_locations ON location_id = datacenter
            LEFT JOIN asset_racks ON rack_id = rack
            LEFT JOIN switchports ON assets.id = switchports.asset_id
            LEFT JOIN switchmanager ON switchports.switch = switchmanager.id
            WHERE assets.id = ? OR assets.hostname LIKE ?`;
        const params = [isNumeric ? parseInt(query, 10) : query, `%${ query }%`];

        try {
            const [rows] = await db.query(sql, params);
            if (!rows || rows.length === 0) {
                await context.sendActivity(MessageFactory.text(
                    `Unable to find any asset with id '${ query }' or with '${ query }' in the hostname`
                ));
                return;
            }
            const r = rows[0];
            const unit = r.unit_start !== r.unit_end
                ? `${ r.unit_start }-${ r.unit_end }`
                : r.unit_start;
            const text = [
                `**Asset:** ${ r.real_asset_id }`,
                `**Hostname:** ${ r.hostname }`,
                `**Status:** ${ r.status }`,
                `**Rack:** ${ r.location_name } ${ r.rack_name }`,
                `**Unit:** ${ unit }`,
                `**Network:** Switch${ r.name } Port ${ r.port }`,
                `https://my.interserver.net/admin/view_server_order?id=${ r.order_id }`,
                `https://my.interserver.net/admin/asset_form?id=${ r.real_asset_id }`
            ].join('\n');
            await context.sendActivity(MessageFactory.text(text));
        } catch (err) {
            await context.sendActivity(MessageFactory.text(`Error searching assets: ${ err.message }`));
            console.error('Asset search error:', err);
        }
    }
};
