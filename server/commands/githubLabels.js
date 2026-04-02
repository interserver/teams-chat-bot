const { MessageFactory } = require('botbuilder');
const { Octokit } = require('@octokit/rest');

const OWNER = process.env.GITHUB_OWNER || 'interserver';
const REPO = process.env.GITHUB_REPO || 'my';

let octokit;
function getOctokit() {
    if (!octokit) {
        octokit = new Octokit({ auth: process.env.GITHUB_TOKEN });
    }
    return octokit;
}

module.exports = {
    match(text, lcText, { ima }) {
        if (ima !== 'admin') return null;
        let m;
        if (/^(labels|label list|labels list|list labels)$/i.test(lcText)) {
            return { action: 'list' };
        }
        if ((m = text.match(/^label\s+create\s+(\S+)\s+#?([0-9a-fA-F]{6})\s*(.*)$/i))) {
            return { action: 'create', name: m[1], color: m[2], description: (m[3] || '').trim() };
        }
        if ((m = text.match(/^label\s+update\s+(\S+)\s+(\S+)\s+#?([0-9a-fA-F]{6})\s*(.*)$/i))) {
            return { action: 'update', name: m[1], newName: m[2], color: m[3], description: (m[4] || '').trim() };
        }
        if ((m = text.match(/^label\s+add\s+(\d+)\s+(.+)$/i))) {
            return { action: 'add', issueId: parseInt(m[1], 10), label: m[2].trim() };
        }
        if ((m = text.match(/^label\s+remove\s+(\d+)\s+(.+)$/i))) {
            return { action: 'remove', issueId: parseInt(m[1], 10), label: m[2].trim() };
        }
        return null;
    },
    async execute(match, { context }) {
        const gh = getOctokit();
        try {
            switch (match.action) {
            case 'list': {
                const { data: labels } = await gh.issues.listLabelsForRepo({
                    owner: OWNER,
                    repo: REPO,
                    per_page: 100
                });
                if (labels.length === 0) {
                    await context.sendActivity(MessageFactory.text('No labels found.'));
                    return;
                }
                let text = `**Labels List** (${ labels.length })\n`;
                for (const label of labels) {
                    text += `- ${ label.name } (Color: #${ label.color })\n`;
                }
                await context.sendActivity(MessageFactory.text(text));
                break;
            }
            case 'create': {
                await gh.issues.createLabel({
                    owner: OWNER,
                    repo: REPO,
                    name: match.name,
                    color: match.color,
                    description: match.description || undefined
                });
                await context.sendActivity(MessageFactory.text(`Label '${ match.name }' created successfully.`));
                break;
            }
            case 'update': {
                await gh.issues.updateLabel({
                    owner: OWNER,
                    repo: REPO,
                    name: match.name,
                    new_name: match.newName,
                    color: match.color,
                    description: match.description || undefined
                });
                await context.sendActivity(MessageFactory.text(`Label '${ match.name }' updated successfully to '${ match.newName }'.`));
                break;
            }
            case 'add': {
                await gh.issues.addLabels({
                    owner: OWNER,
                    repo: REPO,
                    issue_number: match.issueId,
                    labels: [match.label]
                });
                await context.sendActivity(MessageFactory.text(`Label '${ match.label }' added to issue #${ match.issueId } successfully.`));
                break;
            }
            case 'remove': {
                await gh.issues.removeLabel({
                    owner: OWNER,
                    repo: REPO,
                    issue_number: match.issueId,
                    name: match.label
                });
                await context.sendActivity(MessageFactory.text(`Label '${ match.label }' removed from issue #${ match.issueId } successfully.`));
                break;
            }
            }
        } catch (err) {
            await context.sendActivity(MessageFactory.text(`GitHub API error: ${ err.message }`));
            console.error('GitHub Labels command error:', err);
        }
    }
};
