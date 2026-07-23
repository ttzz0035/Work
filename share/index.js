'use strict';

const { execFile } = require('child_process');

function PingClient(device, logger, events) {
    let data = device;
    let connected = false;
    let reading = false;
    let status = 'connect-off';
    let lastRead = 0;
    let values = {};

    function emitStatus(nextStatus) {
        if (status !== nextStatus) {
            status = nextStatus;
            events.emit('device-status:changed', {
                id: data.id,
                status: status
            });
        }
    }

    function ping(callback) {
        const host = String(data.property.address || '').trim();
        const timeout = Number(data.property.timeoutMs || 1000);

        if (!/^[A-Za-z0-9_.:%-]+$/.test(host) || host.startsWith('-')) {
            callback(new Error('Invalid Ping target'));
            return;
        }

        const isWindows = process.platform === 'win32';
        const command = isWindows ? 'ping.exe' : 'ping';
        const args = isWindows
            ? ['-n', '1', '-w', String(timeout), host]
            : ['-n', '-c', '1', '-W', String(Math.max(1, Math.ceil(timeout / 1000))), host];

        execFile(
            command,
            args,
            {
                shell: false,
                windowsHide: true,
                timeout: timeout + 500
            },
            callback
        );
    }

    this.connect = function () {
        connected = true;
        emitStatus('connect-ok');
        logger.info(`'${data.name}' Ping driver connected`, true);
        return Promise.resolve(true);
    };

    this.disconnect = function () {
        connected = false;
        reading = false;
        emitStatus('connect-off');
        return Promise.resolve(true);
    };

    this.polling = function () {
        if (!connected || reading) {
            return;
        }

        reading = true;

        ping(function (error) {
            reading = false;

            if (!connected) {
                return;
            }

            const reachable = !error;
            const timestamp = Date.now();
            const changed = {};

            for (const tagId in data.tags) {
                const previous = values[tagId]?.value;

                values[tagId] = {
                    id: tagId,
                    value: reachable,
                    type: data.tags[tagId].type
                };

                if (previous !== reachable) {
                    changed[tagId] = values[tagId];
                }
            }

            if (reachable) {
                lastRead = timestamp;
                emitStatus('connect-ok');
            } else {
                emitStatus('connect-error');
            }

            events.emit('device-value:changed', {
                id: data.name,
                values: values
            });

            if (error) {
                logger.warn(`'${data.name}' Ping failed: ${error.message}`);
            }
        });
    };

    this.load = function (newData) {
        data = JSON.parse(JSON.stringify(newData));
        values = {};

        for (const tagId in data.tags) {
            values[tagId] = {
                id: tagId,
                value: null,
                type: data.tags[tagId].type
            };
        }
    };

    this.getValues = function () {
        return values;
    };

    this.getValue = function (tagId) {
        return values[tagId]
            ? {
                id: tagId,
                value: values[tagId].value,
                ts: lastRead
            }
            : null;
    };

    this.getStatus = function () {
        return status;
    };

    this.getTagProperty = function (tagId) {
        const tag = data.tags[tagId];

        return tag
            ? {
                id: tagId,
                name: tag.name,
                type: tag.type,
                format: tag.format
            }
            : null;
    };

    this.setValue = function () {
        return Promise.resolve(false);
    };

    this.isConnected = function () {
        return connected;
    };

    this.bindAddDaq = function () {
    };

    this.lastReadTimestamp = function () {
        return lastRead;
    };

    this.load(data);
}

module.exports = {
    create: function (data, logger, events) {
        return new PingClient(data, logger, events);
    }
};