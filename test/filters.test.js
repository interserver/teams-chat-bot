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

describe('shouldSkip — untagged container package versions', () => {
    function pkgEnv(eventType, packageVersion) {
        return {
            extra: {
                event_type: eventType,
                data: {
                    action: 'published',
                    registry_package: { name: 'phlix-hub', package_type: 'CONTAINER', package_version: packageVersion }
                }
            }
        };
    }
    const DIGEST = 'sha256:ef7e3ab209e9b8bc65b5d57e13fa9ea929206ed323c68cf62bd8b6da122cc229';

    it('silently drops a container version whose tag name is empty', () => {
        const env = pkgEnv('registry_package', { name: DIGEST, container_metadata: { tag: { name: '', digest: DIGEST } } });
        assert.equal(shouldSkip(env), '__SILENT__');
    });
    it('drops a whitespace-only tag under the legacy package event too', () => {
        const env = pkgEnv('package', { name: DIGEST, container_metadata: { tag: { name: '  ' } } });
        env.extra.data.package = env.extra.data.registry_package;
        delete env.extra.data.registry_package;
        assert.equal(shouldSkip(env), '__SILENT__');
    });
    it('passes a tagged container version through', () => {
        const env = pkgEnv('registry_package', { name: DIGEST, container_metadata: { tag: { name: 'latest', digest: DIGEST } } });
        assert.equal(shouldSkip(env), null);
    });
    it('passes through when container_metadata is absent (older producer, npm, stripped data)', () => {
        assert.equal(shouldSkip(pkgEnv('registry_package', { name: DIGEST })), null);
        assert.equal(shouldSkip(pkgEnv('registry_package', { name: '2.1.0' })), null);
        assert.equal(shouldSkip({ extra: { event_type: 'registry_package', data: { __stripped: true } } }), null);
        assert.equal(shouldSkip({ extra: { event_type: 'registry_package', data: {} } }), null);
    });
});
