const { MessageFactory } = require('botbuilder');
const axios = require('axios');

const TICKET_URL = process.env.TICKET_API_URL || 'https://mystage.interserver.net/admin/ajax/create_ticket.php';

const VALID_DEPTS = [
    'hardware-modifications', 'developer escalation', 'win-migrations', 'win-escalations',
    'new unassigned', 'host department', 'new features', 'mail baby', 'hdbilling',
    'provisioning', 'migrations', 'level 2', 'hardware', 'billing', 'windows',
    'escalation', 'security', 'general', 'int-1', 'support', 'legal', 'sales',
    'abuse', 'bugs', 'hd'
];

const VALID_DEPT_RE = new RegExp(
    `^add (${ VALID_DEPTS.map(d => d.replace(/[-]/g, '\\$&')).join('|') }) ticket (.*)$`,
    'msi'
);
const INVALID_DEPT_RE = /^add (\w+(?:[ -]\w+)?) ticket /msi;

module.exports = {
    match(text, lcText, { ima }) {
        if (ima !== 'admin') return null;
        const m = lcText.match(VALID_DEPT_RE);
        if (m) return { dept: m[1], msg: m[2] };
        const bad = lcText.match(INVALID_DEPT_RE);
        if (bad) return { invalidDept: bad[1] };
        return null;
    },
    async execute({ dept, msg, invalidDept }, { context, member, email }) {
        if (invalidDept) {
            const deptList = VALID_DEPTS.map(d => `\`${ d }\``).join(', ');
            await context.sendActivity(MessageFactory.text(
                `**"${ invalidDept }"** is not a valid department.\n\nValid departments: ${ deptList }\n\nUsage: \`add <department> ticket <subject>\``
            ));
            return;
        }
        const name = member.name;
        const lines = msg.trim().split(/\r?\n/);
        let subject, body;
        if (lines.length === 1) {
            subject = body = lines[0];
        } else {
            subject = lines[0];
            body = lines.slice(1).join('\n');
        }
        console.log('Subject:', subject);
        console.log('Message:', body);
        try {
            const response = await axios.post(TICKET_URL,
                new URLSearchParams({ subject, body, dept, email, name }),
                { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } });
            if (response.status === 200) {
                await context.sendActivity(MessageFactory.text(response.data));
                console.log('✅ Success:', response.data);
            } else {
                await context.sendActivity(MessageFactory.text(response.data));
                console.log(`⚠️ Unexpected status ${ response.status }:`, response.data);
            }
        } catch (error) {
            if (error.response) {
                await context.sendActivity(MessageFactory.text(error.response.data));
                console.log(`❌ Error ${ error.response.status }:`, error.response.data);
            } else {
                await context.sendActivity(MessageFactory.text(error.message));
                console.log('❌ Request failed:', error.message);
            }
        }
    }
};
