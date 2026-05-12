const { describe, it } = require('node:test');
const assert = require('node:assert/strict');

const { shouldSkip } = require('../server/queue/filters');

function checkRunEnv(name, conclusion = '') {
    return {
        extra: {
            event_type: 'check_run',
            data: { check_run: { name, conclusion, status: 'completed' } }
        }
    };
}

function workflowJobEnv(name, conclusion = '') {
    return {
        extra: {
            event_type: 'workflow_job',
            data: { workflow_job: { name, workflow_name: name, conclusion, status: 'completed' } }
        }
    };
}

describe('shouldSkip — unexpanded `${{ matrix.* }}` placeholders', () => {
    it('passes through a real check_run name', () => {
        assert.equal(shouldSkip(checkRunEnv('Test · PHP 8.3 · candy-core', 'success')), null);
    });
    it('silently skips check_runs with literal ${{ matrix.lib }} in the name', () => {
        // GH Actions emits these when a matrix job is skipped via
        // needs-failure before its matrix can expand. Pure noise.
        assert.equal(shouldSkip(checkRunEnv('Coverage · ${{ matrix.lib }}', 'skipped')), '__SILENT__');
        assert.equal(shouldSkip(checkRunEnv('PHPStan ${{ matrix.php }} · ${{ matrix.lib }}', 'skipped')), '__SILENT__');
    });
    it('silently skips workflow_jobs with literal ${{ matrix.* }} in the name', () => {
        assert.equal(shouldSkip(workflowJobEnv('Coverage · ${{ matrix.lib }}', 'skipped')), '__SILENT__');
    });
    it('does not match a check_run that merely mentions the word matrix', () => {
        assert.equal(shouldSkip(checkRunEnv('matrix sanity test', 'success')), null);
    });
});
