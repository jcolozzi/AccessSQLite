// @ts-check
/**
 * AxLite — JS <-> Access Edge Browser Control Bridge
 * Adapted from Access Designer's proven OutboxTransport + chunked receivePart pattern:
 *   - JS queues envelopes in OutboxTransport; VBA polls AxLite.bridge.peek() via RetrieveJavascriptValue
 *   - VBA acks processed sequence numbers via ExecuteJavascript("AxLite.bridge.ack(seq)")
 *   - VBA sends responses via ExecuteJavascript("AxLite.bridge.receive(...)") or chunked receivePart(...)
 */
(function (AxLite) {
    'use strict';

    class BridgeError extends Error {
        /** @param {string} code @param {string} message */
        constructor(code, message) {
            super(message);
            this.name = 'BridgeError';
            this.code = code;
        }
    }

    class EventBus {
        #handlers = new Map();

        on(name, handler) {
            if (!this.#handlers.has(name)) this.#handlers.set(name, new Set());
            this.#handlers.get(name).add(handler);
            return () => this.#handlers.get(name)?.delete(handler);
        }

        emit(name, data) {
            for (const fn of this.#handlers.get(name) || []) {
                try { fn(data); } catch (err) { console.error(`Event "${name}" handler error:`, err); }
            }
        }
    }

    class OutboxTransport {
        #items = [];
        #seq = 0;

        send(envelope) {
            this.#items.push({ seq: ++this.#seq, env: envelope });
        }

        peek() {
            return this.#items.length ? encodeURIComponent(JSON.stringify(this.#items)) : '';
        }

        ack(seq) {
            this.#items = this.#items.filter((item) => item.seq > seq);
        }

        get pending() {
            return this.#items.length;
        }
    }

    class Bridge {
        static TIMEOUT_MS = 45000;

        #transport;
        #events;
        #nextId = 1;
        #pending = new Map();
        #parts = new Map();

        connected = false;
        lastTrace = null;

        constructor(transport, events) {
            this.#transport = transport;
            this.#events = events;
        }

        /**
         * @param {string} action
         * @param {Record<string, any>} [payload]
         * @param {{ timeoutMs?: number }} [options]
         * @returns {Promise<any>}
         */
        call(action, payload = {}, options = {}) {
            const timeoutMs = options.timeoutMs ?? Bridge.TIMEOUT_MS;
            return new Promise((resolve, reject) => {
                const id = this.#nextId++;
                const timer = setTimeout(() => {
                    this.#pending.delete(id);
                    reject(new BridgeError('Timeout', `No reply from Access VBA host for "${action}".`));
                }, timeoutMs);
                this.#pending.set(id, { resolve, reject, timer, action });
                this.#transport.send({ id, action, payload });
            });
        }

        peek() {
            return this.#transport.peek();
        }

        ack(seq) {
            this.#transport.ack(seq);
        }

        receivePart(key, index, count, text) {
            const k = String(key);
            const parts = this.#parts.get(k) ?? new Array(count);
            parts[index] = text;
            this.#parts.set(k, parts);
            for (let i = 0; i < count; i++) {
                if (parts[i] === undefined) return;
            }
            this.#parts.delete(k);
            this.receive(JSON.parse(parts.join('')));
        }

        receive(message) {
            if (!message || typeof message !== 'object') return;
            if (!this.connected) {
                this.connected = true;
                this.#events.emit('bridge.connected', true);
            }
            if (message.event) {
                this.#events.emit(message.event, message.data);
                return;
            }
            const entry = this.#pending.get(message.id);
            if (!entry) return;
            this.#pending.delete(message.id);
            clearTimeout(entry.timer);
            if (message.ok) {
                entry.resolve(message.data);
            } else {
                const code = (message.error && message.error.code) ? message.error.code : 'SQLiteError';
                const msg = (message.error && message.error.message) ? message.error.message : 'Operation failed.';
                entry.reject(new BridgeError(code, msg));
            }
        }
    }

    // Suppress harmless Monaco ResizeObserver frame deferral notices; display genuine boot errors
    window.addEventListener('error', (e) => {
        if (String(e.message || '').startsWith('ResizeObserver loop')) {
            e.stopImmediatePropagation();
            return;
        }
        if (!AxLite.app) {
            const root = document.getElementById('ss-app');
            if (root) root.textContent = `Boot Error: ${e.message} (${e.filename || ''}:${e.lineno || 0})`;
        }
    });
    window.addEventListener('unhandledrejection', (e) => {
        if (!AxLite.app) {
            const root = document.getElementById('ss-app');
            if (root) root.textContent = `Unhandled Boot Promise: ${e.reason?.message || e.reason}`;
        }
    });

    const events = new EventBus();
    const transport = new OutboxTransport();
    const bridge = new Bridge(transport, events);

    AxLite.events = events;
    AxLite.bridge = bridge;
    AxLite.BridgeError = BridgeError;
})(window.AxLite = window.AxLite || {});
