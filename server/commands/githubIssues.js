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
        if (/^(issues|issue list|list issues|issues list)$/i.test(lcText)) {
            return { action: 'list' };
        }
        if ((m = text.match(/^issues?\s*(?:show|view|get|display)\s+(\d+)$/i)) ||
            (m = text.match(/^(?:show|view|get|display)\s+issues?\s+(\d+)$/i))) {
            return { action: 'show', id: parseInt(m[1], 10) };
        }
        if ((m = text.match(/^issues?\s+close\s+(\d+)\s*(.*)$/i))) {
            return { action: 'close', id: parseInt(m[1], 10), comment: (m[2] || '').trim() };
        }
        if ((m = text.match(/^issues?\s+comment\s+(\d+)\s+(.+)$/i))) {
            return { action: 'comment', id: parseInt(m[1], 10), comment: m[2].trim() };
        }
        if ((m = text.match(/^issues?\s+create\s+(.+)$/is))) {
            return { action: 'create', raw: m[1].trim() };
        }
        if (/^(gh|github)\s+help$/i.test(lcText)) {
            return { action: 'help' };
        }
        return null;
    },
    async execute(match, { context }) {
        const gh = getOctokit();
        try {
            switch (match.action) {
            case 'list': {
                const { data: issues } = await gh.issues.listForRepo({
                    owner: OWNER,
                    repo: REPO,
                    state: 'open',
                    per_page: 30
                });
                if (issues.length === 0) {
                    await context.sendActivity(MessageFactory.text('No open issues found.'));
                    return;
                }
                let text = `**Issues List** (${ issues.length })\n`;
                for (const issue of issues) {
                    text += `[**#${ issue.number }**](${ issue.html_url }) ${ issue.title } by [${ issue.user.login }](${ issue.user.html_url })\n`;
                }
                await context.sendActivity(MessageFactory.text(text));
                break;
            }
            case 'show': {
                const [issueRes, commentsRes, labelsRes] = await Promise.all([
                    gh.issues.get({ owner: OWNER, repo: REPO, issue_number: match.id }),
                    gh.issues.listComments({ owner: OWNER, repo: REPO, issue_number: match.id }),
                    gh.issues.listLabelsOnIssue({ owner: OWNER, repo: REPO, issue_number: match.id })
                ]);
                const issue = issueRes.data;
                const comments = commentsRes.data;
                const labels = labelsRes.data;
                let text = `**Issue #${ issue.number } ${ issue.title }**\n`;
                text += `${ issue.body || '(no description)' } by [${ issue.user.login }](${ issue.user.html_url }) on ${ issue.updated_at }\n`;
                if (labels.length > 0) {
                    text += `**Labels:** ${ labels.map(l => l.name).join(', ') }\n`;
                }
                if (comments.length > 0) {
                    text += '**Comments**\n';
                    for (const c of comments) {
                        text += `[${ c.user.login }](${ c.user.html_url }) said: ${ c.body } on ${ c.updated_at }\n`;
                    }
                }
                await context.sendActivity(MessageFactory.text(text));
                break;
            }
            case 'close': {
                if (match.comment) {
                    await gh.issues.createComment({
                        owner: OWNER,
                        repo: REPO,
                        issue_number: match.id,
                        body: match.comment
                    });
                }
                await gh.issues.update({
                    owner: OWNER,
                    repo: REPO,
                    issue_number: match.id,
                    state: 'closed'
                });
                await context.sendActivity(MessageFactory.text(`Issue #${ match.id } closed successfully.`));
                break;
            }
            case 'comment': {
                await gh.issues.createComment({
                    owner: OWNER,
                    repo: REPO,
                    issue_number: match.id,
                    body: match.comment
                });
                await context.sendActivity(MessageFactory.text(`Comment added to issue #${ match.id } successfully.`));
                break;
            }
            case 'create': {
                const lines = match.raw.split(/\r?\n/);
                const title = lines[0];
                const body = lines.length > 1 ? lines.slice(1).join('\n') : '';
                const { data: issue } = await gh.issues.create({
                    owner: OWNER,
                    repo: REPO,
                    title,
                    body
                });
                await context.sendActivity(MessageFactory.text(`Issue created successfully. Issue #${ issue.number }: ${ issue.title }`));
                break;
            }
            case 'help': {
                const cmds = {
                    'issues list': 'List all open issues',
                    'issues show <id>': 'Show details of a specific issue',
                    'issues close <id> [comment]': 'Close an issue with an optional comment',
                    'issues comment <id> <comment>': 'Add a comment to an issue',
                    'issues create <title> [body]': 'Create a new issue',
                    'labels list': 'List all labels',
                    'label create <name> <color> [description]': 'Create a new label',
                    'label update <name> <new_name> <color> [description]': 'Update an existing label',
                    'label add <issue_id> <label>': 'Add a label to an issue',
                    'label remove <issue_id> <label>': 'Remove a label from an issue',
                    'github help': 'Show all available GitHub commands'
                };
                let text = '**GitHub Commands Help**\n';
                for (const [cmd, desc] of Object.entries(cmds)) {
                    text += `\`${ cmd }\` - ${ desc }\n`;
                }
                await context.sendActivity(MessageFactory.text(text));
                break;
            }
            }
        } catch (err) {
            await context.sendActivity(MessageFactory.text(`GitHub API error: ${ err.message }`));
            console.error('GitHub Issues command error:', err);
        }
    }
};
