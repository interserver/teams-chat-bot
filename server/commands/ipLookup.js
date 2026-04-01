module.exports = {
    match(text, lcText, { ima }) {
        if (ima !== 'admin') return null;
        const m = text.match(/.*(where|lookup|query|find|locate|search).*?[^\d](\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})[^\d]?.*/i);
        return m ? { ip: m[2] } : null;
    },
    async execute({ ip }, { context, db }) {
        const [rows] = await db.query(
            `SELECT *, assets.id AS real_asset_id
            FROM ips
            LEFT JOIN vlans ON ips_vlan=vlans_id
            LEFT JOIN switchports ON FIND_IN_SET(ips_vlan, vlans) != 0
            LEFT JOIN assets ON switchports.asset_id=assets.id
            LEFT JOIN asset_types ON type_id=asset_types.asset_id
            LEFT JOIN asset_locations ON location_id=datacenter
            LEFT JOIN asset_racks ON rack_id=rack
            LEFT JOIN switchmanager ON switchports.switch=switchmanager.id
            WHERE ips_ip = ?`,
            [ip]
        );

        if (!rows || rows.length === 0) {
            await context.sendActivity(`Unable to find ${ ip } in our IP database`);
        } else {
            const r = rows[0];
            const unit = r.unit_start !== r.unit_end ? `${ r.unit_start }-${ r.unit_end }` : r.unit_start;
            await context.sendActivity(`Asset: ${ r.real_asset_id }\n
Hostname: ${ r.hostname }\n
Status: ${ r.status }\n
Rack: ${ r.location_name } ${ r.rack_name }\n
Unit: ${ unit }\n
Network: Switch${ r.name } Port ${ r.port }\n
https://my.interserver.net/admin/view_server_order?id=${ r.order_id }\n
https://my.interserver.net/admin/asset_form?id=${ r.real_asset_id }`
            );
        }
    }
};
