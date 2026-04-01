module.exports = {
    match(text, lcText, { ima }) {
        if (ima !== 'admin') return null;
        let m;
        if ((m = lcText.match(/^(mark|set) (\S+) (1|0|unavailable|available|enabled|disabled|disable|enable|off|on|usable|unusable)$/i))) {
            const value = m[3].toLowerCase();
            const valuesOn = ['1', 'available', 'enabled', 'enable', 'on', 'usable'];
            const resolved = valuesOn.includes(value) ? '1' : '0';
            return { server: m[2], field: 'available', value: resolved };
        }
        if ((m = lcText.match(/^(mark|set) (\S+) (\S+)( | *= *| *to *)(\S+)$/i))) {
            return { server: m[2], field: m[3], value: m[5] };
        }
        if ((m = lcText.match(/^(mark|set) (\S+) on (\S+)( | *= *| *to *)(\S+)$/i))) {
            return { server: m[3], field: m[2], value: m[5] };
        }
        return null;
    },
    async execute({ server, field, value }, { context, bot }) {
        await bot.setMaster(context, server, field, value);
    }
};
