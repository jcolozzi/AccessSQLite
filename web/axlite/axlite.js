// @ts-check
/**
 * AxLite — Monaco-powered SQLite Front-End Client for Microsoft Access
 * Features:
 *   - Multi-tab Monaco SQL editor with live SQLite schema IntelliSense
 *   - Schema Explorer (Tables, Columns w/ PK/FK/NOT NULL/Type badges, Indexes, Views, Triggers)
 *   - Stored Queries manager (persisted in _axlite_saved_queries)
 *   - Executes all SQL operations (SELECT, DML, DDL, PRAGMA, RETURNING, transactions, multi-statement scripts)
 *   - Sortable Results Grid, Execution Messages, Explain Query Plan, CSV/JSON Export & Query History
 */
(function (AxLite) {
    'use strict';

    const bridge = AxLite.bridge;

    const SQLITE_KEYWORDS_AND_FUNCS = [
        'PRAGMA journal_mode;', 'PRAGMA foreign_keys;', 'PRAGMA table_xinfo(', 'PRAGMA index_list(',
        'PRAGMA integrity_check;', 'PRAGMA compile_options;', 'EXPLAIN QUERY PLAN ',
        'BEGIN IMMEDIATE TRANSACTION;', 'COMMIT;', 'ROLLBACK;', 'SAVEPOINT ', 'VACUUM;',
        'COALESCE(', 'IFNULL(', 'NULLIF(', 'IIF(', 'PRINTF(', 'SUBSTR(', 'INSTR(', 'REPLACE(',
        'ROUND(', 'ABS(', 'TOTAL(', 'GROUP_CONCAT(', 'JSON_EXTRACT(', 'JSON_OBJECT(', 'JSON_GROUP_ARRAY(',
        'ROW_NUMBER() OVER (', 'LAG(', 'LEAD(', 'RETURNING *'
    ];

    /** Loads Monaco 0.52.2 AMD build using same-origin workerMain.js (safe on https://msaccess/). */
    function loadMonaco() {
        const base = new URL('../vendor/monaco/', location.href).href;
        window['MonacoEnvironment'] = {
            getWorkerUrl: (_moduleId, label) => `${base}vs/base/worker/workerMain.js#${label}`
        };
        return new Promise((resolve, reject) => {
            const amdRequire = window['require'];
            if (!amdRequire) {
                reject(new Error('Monaco loader.js was not found at ../vendor/monaco/vs/loader.js'));
                return;
            }
            amdRequire.config({ paths: { vs: `${base}vs` } });
            amdRequire(['vs/editor/editor.main'], () => resolve(window['monaco']), reject);
        });
    }

    function quoteIdent(name) {
        return `"${String(name).replace(/"/g, '""')}"`;
    }

    function formatBytes(bytes) {
        const n = Number(bytes || 0);
        if (n <= 0) return 'in-memory';
        if (n < 1024) return `${n} B`;
        if (n < 1048576) return `${(n / 1024).toFixed(1)} KB`;
        return `${(n / 1048576).toFixed(2)} MB`;
    }

    function el(tag, attrs = {}, ...children) {
        const node = document.createElement(tag);
        for (const [k, v] of Object.entries(attrs || {})) {
            if (v === null || v === undefined || v === false) continue;
            if (k === 'class') node.className = v;
            else if (k === 'title') node.title = v;
            else if (k === 'disabled') node.disabled = Boolean(v);
            else if (k === 'value') /** @type {any} */ (node).value = v;
            else if (k === 'style' && typeof v === 'object') Object.assign(node.style, v);
            else if (k === 'on' && typeof v === 'object') {
                for (const [ev, fn] of Object.entries(v)) node.addEventListener(ev, fn);
            } else {
                node.setAttribute(k, String(v));
            }
        }
        for (const child of children.flat()) {
            if (child === null || child === undefined || child === false) continue;
            node.append(child instanceof Node ? child : document.createTextNode(String(child)));
        }
        return node;
    }

    function setChildren(target, ...children) {
        const clean = [];
        for (const child of children.flat()) {
            if (child === null || child === undefined || child === false) continue;
            clean.push(child instanceof Node ? child : document.createTextNode(String(child)));
        }
        target.replaceChildren(...clean);
    }

    class AxLiteApp {
        #monaco;
        #editor;
        #root;
        #darkMode = true;
        #busy = false;

        #dbInfo = {
            dbPath: '',
            dbName: '(Connecting...)',
            engineVersion: '',
            driverTitle: '',
            journalMode: '',
            foreignKeys: 1,
            fileSizeBytes: 0,
            inTransaction: false
        };

        #schema = { tables: [], views: [], triggers: [] };
        #savedQueries = [];
        #expandedNodes = new Set();
        #treeFilter = '';

        /** @type {Array<{ id: string, title: string, savedName: string, category: string, description: string, model: any, savedAltVer: number, viewState: any }>} */
        #tabs = [];
        #activeTabId = '';
        #tabSeq = 0;

        #rowLimit = 1000;
        #activeBottomTab = 'results'; // results | messages | plan | history
        #lastResult = null;
        #sortColIdx = -1;
        #sortAsc = true;
        #messages = [];
        #planNodes = [];
        #history = [];

        // DOM references
        #topbarEl;
        #treeEl;
        #queryTabsEl;
        #editorPaneEl;
        #monacoHostEl;
        #resBarEl;
        #resBodyEl;
        #statusEl;

        constructor(rootEl, monaco) {
            this.#root = rootEl;
            this.#monaco = monaco;
            this.#buildShell();
            this.#initMonaco();
            this.#registerSqlCompletionProvider();
            this.#wireSaveModal();
            this.#createTab('Query 1', '-- Press Ctrl+Enter or F5 to execute SQL (runs selection if highlighted)\nSELECT sqlite_version() AS SQLiteVersion;\n');
            this.connectInitial();
        }

        #buildShell() {
            this.#topbarEl = el('header', { class: 'ss-topbar' });

            const searchInput = el('input', {
                class: 'ss-input',
                type: 'search',
                placeholder: 'Filter tables, columns, queries...',
                on: {
                    input: (e) => {
                        this.#treeFilter = String(e.target.value || '').trim().toLowerCase();
                        this.#renderTree();
                    }
                }
            });
            this.#treeEl = el('div', { class: 'ss-tree' });
            const sidebar = el('aside', { class: 'ss-sidebar' },
                el('div', { class: 'ss-sidebar-search' }, searchInput),
                this.#treeEl
            );

            this.#queryTabsEl = el('div', { class: 'ss-query-tabs' });
            this.#monacoHostEl = el('div', { class: 'ss-monaco-host' });
            this.#editorPaneEl = el('div', { class: 'ss-editor-pane' }, this.#monacoHostEl);

            const splitter = el('div', { class: 'ss-splitter', title: 'Drag to resize editor / results' });
            this.#wireSplitter(splitter);

            this.#resBarEl = el('div', { class: 'ss-res-bar' });
            this.#resBodyEl = el('div', { class: 'ss-res-body' });
            const resultsPane = el('div', { class: 'ss-results-pane' }, this.#resBarEl, this.#resBodyEl);

            const center = el('section', { class: 'ss-center' },
                this.#queryTabsEl,
                this.#editorPaneEl,
                splitter,
                resultsPane
            );

            this.#statusEl = el('footer', { class: 'ss-statusbar' });

            this.#root.replaceChildren(
                this.#topbarEl,
                el('div', { class: 'ss-workspace' }, sidebar, center),
                this.#statusEl
            );
        }

        #wireSplitter(splitter) {
            splitter.addEventListener('mousedown', (startEv) => {
                startEv.preventDefault();
                const startY = startEv.clientY;
                const startHeight = this.#editorPaneEl.getBoundingClientRect().height;
                const parentHeight = this.#editorPaneEl.parentElement.getBoundingClientRect().height;

                const onMove = (moveEv) => {
                    const next = Math.max(110, Math.min(parentHeight - 130, startHeight + (moveEv.clientY - startY)));
                    this.#editorPaneEl.style.height = `${next}px`;
                };
                const onUp = () => {
                    window.removeEventListener('mousemove', onMove);
                    window.removeEventListener('mouseup', onUp);
                };
                window.addEventListener('mousemove', onMove);
                window.addEventListener('mouseup', onUp);
            });
        }

        #initMonaco() {
            this.#editor = this.#monaco.editor.create(this.#monacoHostEl, {
                theme: 'vs-dark',
                language: 'sql',
                automaticLayout: true,
                fontSize: 13.5,
                fontFamily: '"Cascadia Code", Consolas, "Courier New", monospace',
                minimap: { enabled: false },
                tabSize: 4,
                insertSpaces: true,
                scrollBeyondLastLine: false,
                wordWrap: 'on',
                suggestOnTriggerCharacters: true,
                model: null
            });

            // Keyboard shortcuts: Ctrl+Enter or F5 -> Run SQL; Ctrl+S -> Save Stored Query
            this.#editor.addCommand(this.#monaco.KeyMod.CtrlCmd | this.#monaco.KeyCode.Enter, () => this.runActiveSql());
            this.#editor.addCommand(this.#monaco.KeyCode.F5, () => this.runActiveSql());
            this.#editor.addCommand(this.#monaco.KeyMod.CtrlCmd | this.#monaco.KeyCode.KeyS, () => this.promptSaveQuery());
            this.#editor.onDidChangeCursorPosition(() => this.#renderStatusBar());

            document.addEventListener('keydown', (e) => {
                if ((e.ctrlKey || e.metaKey) && String(e.key || '').toLowerCase() === 's') {
                    e.preventDefault();
                    this.promptSaveQuery();
                } else if (e.key === 'F5') {
                    e.preventDefault();
                    this.runActiveSql();
                }
            });
        }

        #registerSqlCompletionProvider() {
            const monaco = this.#monaco;
            monaco.languages.registerCompletionItemProvider('sql', {
                triggerCharacters: ['.', ' ', '"'],
                provideCompletionItems: (model, position) => {
                    const word = model.getWordUntilPosition(position);
                    const range = {
                        startLineNumber: position.lineNumber,
                        endLineNumber: position.lineNumber,
                        startColumn: word.startColumn,
                        endColumn: word.endColumn
                    };

                    const linePrefix = model.getValueInRange({
                        startLineNumber: position.lineNumber,
                        startColumn: 1,
                        endLineNumber: position.lineNumber,
                        endColumn: position.column
                    });

                    const dotMatch = /([A-Za-z0-9_"]+)\.\s*[A-Za-z0-9_]*$/.exec(linePrefix);
                    const suggestions = [];

                    if (dotMatch) {
                        const rawTable = dotMatch[1].replace(/"/g, '').toLowerCase();
                        const matchObj = [...this.#schema.tables, ...this.#schema.views]
                            .find((t) => t.name.toLowerCase() === rawTable);
                        if (matchObj) {
                            for (const c of matchObj.columns || []) {
                                suggestions.push({
                                    label: c.name,
                                    kind: monaco.languages.CompletionItemKind.Field,
                                    detail: `${matchObj.name}.${c.name} (${c.type || 'ANY'})`,
                                    insertText: /\W/.test(c.name) ? quoteIdent(c.name) : c.name,
                                    range
                                });
                            }
                            return { suggestions };
                        }
                    }

                    for (const tbl of this.#schema.tables || []) {
                        suggestions.push({
                            label: tbl.name,
                            kind: monaco.languages.CompletionItemKind.Class,
                            detail: `Table (${tbl.rowCount ?? 0} rows)`,
                            insertText: /\W/.test(tbl.name) ? quoteIdent(tbl.name) : tbl.name,
                            range
                        });
                        for (const col of tbl.columns || []) {
                            suggestions.push({
                                label: col.name,
                                kind: monaco.languages.CompletionItemKind.Field,
                                detail: `${tbl.name} · ${col.type || 'ANY'}${col.pk ? ' [PK]' : ''}`,
                                insertText: /\W/.test(col.name) ? quoteIdent(col.name) : col.name,
                                range
                            });
                        }
                    }

                    for (const vw of this.#schema.views || []) {
                        suggestions.push({
                            label: vw.name,
                            kind: monaco.languages.CompletionItemKind.Interface,
                            detail: 'View',
                            insertText: /\W/.test(vw.name) ? quoteIdent(vw.name) : vw.name,
                            range
                        });
                    }

                    for (const kw of SQLITE_KEYWORDS_AND_FUNCS) {
                        suggestions.push({
                            label: kw,
                            kind: monaco.languages.CompletionItemKind.Function,
                            detail: 'SQLite built-in / PRAGMA',
                            insertText: kw,
                            range
                        });
                    }

                    return { suggestions };
                }
            });
        }

        // =====================================================================
        // Connection & Schema Sync
        // =====================================================================

        async connectInitial() {
            this.#renderAll();
            try {
                const initData = await bridge.call('axlite.init');
                this.#applyHostState(initData);
                this.#logMessage('ok', `Connected to ${this.#dbInfo.dbPath || ':memory:'} (${this.#dbInfo.driverTitle} v${this.#dbInfo.engineVersion})`);
                if (this.#schema.tables.length > 0 && this.#tabs.length === 1) {
                    const preferred = this.#schema.tables.find((t) => !String(t.name).startsWith('_')) || this.#schema.tables[0];
                    const firstTbl = preferred.name;
                    this.#expandedNodes.add(`tbl:${firstTbl}`);
                    const tab = this.activeTab;
                    if (tab && tab.model.getValue().includes('SELECT sqlite_version()')) {
                        const starterSql = `-- Connected to ${this.#dbInfo.dbName}\n-- Press Ctrl+Enter or F5 to execute; highlight any block to run just that selection.\nSELECT * FROM ${quoteIdent(firstTbl)} LIMIT 100;\n`;
                        tab.model.setValue(starterSql);
                        tab.savedAltVer = tab.model.getAlternativeVersionId();
                        this.#renderAll();
                        await this.runActiveSql(starterSql);
                        return;
                    }
                }
                this.#renderAll();
            } catch (err) {
                this.#logMessage('err', `Host connection error: ${err.message}`);
                this.#activeBottomTab = 'messages';
                this.#renderAll();
            }
        }

        #applyHostState(data) {
            if (!data) return;
            if (data.dbPath !== undefined) this.#dbInfo.dbPath = data.dbPath;
            if (data.dbName !== undefined) this.#dbInfo.dbName = data.dbName;
            if (data.engineVersion !== undefined) this.#dbInfo.engineVersion = data.engineVersion;
            if (data.driverTitle !== undefined) this.#dbInfo.driverTitle = data.driverTitle;
            if (data.journalMode !== undefined) this.#dbInfo.journalMode = data.journalMode;
            if (data.foreignKeys !== undefined) this.#dbInfo.foreignKeys = data.foreignKeys;
            if (data.fileSizeBytes !== undefined) this.#dbInfo.fileSizeBytes = data.fileSizeBytes;
            if (data.inTransaction !== undefined) this.#dbInfo.inTransaction = Boolean(data.inTransaction);
            if (data.schema) this.#schema = data.schema;
            if (data.savedQueries) this.#savedQueries = data.savedQueries;
        }

        async refreshSchema() {
            try {
                const res = await bridge.call('axlite.init');
                this.#applyHostState(res);
                this.#renderAll();
            } catch (err) {
                this.#logMessage('err', `Schema refresh failed: ${err.message}`);
                this.#activeBottomTab = 'messages';
                this.#renderResultsPane();
            }
        }

        async openDatabaseDialog(mode) {
            try {
                const res = await bridge.call('axlite.openDialog', { mode });
                if (res && !res.cancelled) {
                    this.#expandedNodes.clear();
                    this.#lastResult = null;
                    this.#applyHostState(res);
                    if (this.#schema.tables.length > 0) {
                        this.#expandedNodes.add(`tbl:${this.#schema.tables[0].name}`);
                    }
                    this.#logMessage('ok', `Opened database: ${this.#dbInfo.dbPath}`);
                    this.#renderAll();
                }
            } catch (err) {
                this.#logMessage('err', `Open database error: ${err.message}`);
                this.#activeBottomTab = 'messages';
                this.#renderAll();
            }
        }

        async openInMemoryDb() {
            try {
                const res = await bridge.call('axlite.openPath', { dbPath: ':memory:' });
                this.#expandedNodes.clear();
                this.#lastResult = null;
                this.#applyHostState(res);
                this.#logMessage('ok', 'Switched to isolated :memory: SQLite database.');
                this.#renderAll();
            } catch (err) {
                this.#logMessage('err', err.message);
                this.#activeBottomTab = 'messages';
                this.#renderAll();
            }
        }

        // =====================================================================
        // SQL Execution, Explain & DB Maintenance
        // =====================================================================

        get activeTab() {
            return this.#tabs.find((t) => t.id === this.#activeTabId) || null;
        }

        #getSqlToExecute() {
            const model = this.#editor.getModel();
            if (!model) return '';
            const sel = this.#editor.getSelection();
            if (sel && !sel.isEmpty()) {
                const selectedText = model.getValueInRange(sel).trim();
                if (selectedText.length > 0) return selectedText;
            }
            return model.getValue().trim();
        }

        async runActiveSql(overrideSql = '') {
            const sql = (overrideSql || this.#getSqlToExecute()).trim();
            if (!sql || this.#busy) return;

            this.#busy = true;
            this.#renderTopbar();
            this.#renderStatusBar();

            const startedAt = new Date().toLocaleTimeString();
            try {
                const res = await bridge.call('axlite.execSql', {
                    sql,
                    rowLimit: this.#rowLimit
                });

                this.#dbInfo.inTransaction = Boolean(res.inTransaction);
                if (res.fileSizeBytes !== undefined) this.#dbInfo.fileSizeBytes = res.fileSizeBytes;

                this.#history.unshift({
                    time: startedAt,
                    sql,
                    ok: true,
                    elapsedMs: res.totalElapsedMs,
                    summary: res.hasResultSet
                        ? `${res.rowCount} row(s)${res.truncated ? '+' : ''}`
                        : `${res.totalChanges} row(s) affected`
                });

                const stmtLines = (res.statements || []).map((s) =>
                    s.isQuery
                        ? `[#${s.index}] Query returned ${s.rowCount} row(s) in ${Number(s.elapsedMs).toFixed(2)} ms — ${s.sqlPreview}`
                        : `[#${s.index}] ${s.rowsAffected} row(s) affected (lastRowId=${s.lastInsertRowId}) in ${Number(s.elapsedMs).toFixed(2)} ms — ${s.sqlPreview}`
                );
                this.#logMessage('ok',
                    `Executed ${(res.statements || []).length} statement(s) in ${Number(res.totalElapsedMs).toFixed(2)} ms.` +
                    (stmtLines.length ? '\n' + stmtLines.join('\n') : '')
                );

                if (res.hasResultSet) {
                    this.#lastResult = res;
                    this.#sortColIdx = -1;
                    this.#sortAsc = true;
                    this.#activeBottomTab = 'results';
                } else {
                    this.#lastResult = null;
                    this.#activeBottomTab = 'messages';
                }

                if (res.schemaChanged) {
                    const refreshed = await bridge.call('axlite.init');
                    this.#applyHostState(refreshed);
                }
            } catch (err) {
                this.#history.unshift({
                    time: startedAt,
                    sql,
                    ok: false,
                    elapsedMs: 0,
                    summary: err.message
                });
                this.#logMessage('err', err.message);
                this.#activeBottomTab = 'messages';
            } finally {
                this.#busy = false;
                this.#renderAll();
                this.#editor.focus();
            }
        }

        async explainActiveSql() {
            const sql = this.#getSqlToExecute();
            if (!sql || this.#busy) return;
            try {
                const res = await bridge.call('axlite.explainSql', { sql });
                this.#planNodes = res.nodes || [];
                this.#activeBottomTab = 'plan';
                this.#renderResultsPane();
            } catch (err) {
                this.#logMessage('err', `Explain Plan failed: ${err.message}`);
                this.#activeBottomTab = 'messages';
                this.#renderResultsPane();
            }
        }

        async runDbMaintenance(kind) {
            if (!kind || this.#busy) return;
            try {
                const res = await bridge.call('axlite.dbAction', { kind });
                if (res && !res.cancelled) {
                    if (res.fileSizeBytes !== undefined) this.#dbInfo.fileSizeBytes = res.fileSizeBytes;
                    this.#logMessage(res.ok === false ? 'err' : 'ok', res.message || `${kind} completed.`);
                    this.#activeBottomTab = 'messages';
                    await this.refreshSchema();
                }
            } catch (err) {
                this.#logMessage('err', `${kind} error: ${err.message}`);
                this.#activeBottomTab = 'messages';
                this.#renderResultsPane();
            }
        }

        async exportResultsCsv() {
            if (!this.#lastResult || !this.#lastResult.columns || !this.#lastResult.columns.length) return;
            const cols = this.#lastResult.columns;
            const rows = this.#getSortedRows();
            const esc = (v) => {
                if (v === null || v === undefined) return '';
                const s = String(v);
                return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
            };
            const lines = [cols.map((c) => esc(c.name)).join(',')];
            for (const r of rows) {
                lines.push(r.map((cell) => esc(cell)).join(','));
            }
            const csvText = lines.join('\r\n');
            try {
                const res = await bridge.call('axlite.dbAction', {
                    kind: 'export_csv',
                    csvText,
                    suggestedName: `${(this.activeTab?.title || 'query_results').replace(/[^A-Za-z0-9_-]/g, '_')}.csv`
                });
                if (res && !res.cancelled) {
                    this.#logMessage('ok', res.message || 'Exported CSV successfully.');
                }
            } catch (err) {
                this.#logMessage('err', `CSV export failed: ${err.message}`);
                this.#activeBottomTab = 'messages';
                this.#renderResultsPane();
            }
        }

        async copyResultsJson() {
            if (!this.#lastResult || !this.#lastResult.columns) return;
            const cols = this.#lastResult.columns;
            const rows = this.#getSortedRows();
            const objs = rows.map((r) => {
                const o = {};
                cols.forEach((c, idx) => { o[c.name] = r[idx]; });
                return o;
            });
            const jsonText = JSON.stringify(objs, null, 2);
            try {
                await navigator.clipboard.writeText(jsonText);
                this.#logMessage('ok', `Copied ${objs.length} row(s) as JSON to clipboard.`);
                this.#renderStatusBar(`Copied ${objs.length} row(s) as JSON`);
            } catch (_err) {
                this.openSqlInNewTab('Results.json', jsonText);
            }
        }

        // =====================================================================
        // Editor Tabs & Stored Queries
        // =====================================================================

        #createTab(title, initialSql = '', meta = {}) {
            if (this.activeTab) {
                this.activeTab.viewState = this.#editor.saveViewState();
            }
            const id = `tab_${++this.#tabSeq}`;
            const model = this.#monaco.editor.createModel(initialSql, 'sql');
            model.onDidChangeContent(() => {
                this.#renderQueryTabs();
            });
            const tab = {
                id,
                title,
                savedName: meta.savedName || '',
                category: meta.category || 'General',
                description: meta.description || '',
                model,
                savedAltVer: model.getAlternativeVersionId(),
                viewState: null
            };
            this.#tabs.push(tab);
            this.#switchTab(id);
            return tab;
        }

        openSqlInNewTab(title, sqlText, meta = {}) {
            const existing = meta.savedName
                ? this.#tabs.find((t) => t.savedName && t.savedName.toLowerCase() === meta.savedName.toLowerCase())
                : null;
            if (existing) {
                this.#switchTab(existing.id);
                return existing;
            }
            return this.#createTab(title, sqlText, meta);
        }

        #switchTab(tabId) {
            const current = this.activeTab;
            if (current && current.id !== tabId) {
                current.viewState = this.#editor.saveViewState();
            }
            const next = this.#tabs.find((t) => t.id === tabId);
            if (!next) return;
            this.#activeTabId = next.id;
            this.#editor.setModel(next.model);
            if (next.viewState) this.#editor.restoreViewState(next.viewState);
            this.#editor.focus();
            this.#renderQueryTabs();
            this.#renderStatusBar();
        }

        #closeTab(tabId, ev) {
            if (ev) ev.stopPropagation();
            if (this.#tabs.length <= 1) {
                this.activeTab.model.setValue('');
                this.activeTab.title = 'Query 1';
                this.activeTab.savedName = '';
                this.#renderQueryTabs();
                return;
            }
            const idx = this.#tabs.findIndex((t) => t.id === tabId);
            if (idx < 0) return;
            const [removed] = this.#tabs.splice(idx, 1);
            removed.model.dispose();
            if (this.#activeTabId === tabId) {
                const fallback = this.#tabs[Math.min(idx, this.#tabs.length - 1)];
                this.#switchTab(fallback.id);
            } else {
                this.#renderQueryTabs();
            }
        }

        #insertAtCursor(text) {
            const sel = this.#editor.getSelection();
            if (!sel) return;
            this.#editor.executeEdits('schema-insert', [{
                range: sel,
                text,
                forceMoveMarkers: true
            }]);
            this.#editor.focus();
        }

        #wireSaveModal() {
            const modal = document.getElementById('ss-save-modal');
            const nameIn = /** @type {HTMLInputElement} */ (document.getElementById('ss-q-name'));
            const catIn = /** @type {HTMLInputElement} */ (document.getElementById('ss-q-cat'));
            const descIn = /** @type {HTMLInputElement} */ (document.getElementById('ss-q-desc'));
            const btnCancel = document.getElementById('ss-q-cancel');
            const btnConfirm = document.getElementById('ss-q-confirm');

            const close = () => {
                modal.hidden = true;
                this.#editor.focus();
            };

            btnCancel.addEventListener('click', close);
            modal.addEventListener('click', (e) => { if (e.target === modal) close(); });

            btnConfirm.addEventListener('click', async () => {
                const qName = nameIn.value.trim();
                const tab = this.activeTab;
                if (!qName || !tab) {
                    nameIn.focus();
                    return;
                }
                const sqlText = tab.model.getValue().trim();
                if (!sqlText) return;
                try {
                    const res = await bridge.call('axlite.saveQuery', {
                        name: qName,
                        category: catIn.value.trim() || 'General',
                        description: descIn.value.trim(),
                        sql: sqlText
                    });
                    this.#savedQueries = res.savedQueries || [];
                    tab.title = qName;
                    tab.savedName = qName;
                    tab.category = catIn.value.trim() || 'General';
                    tab.description = descIn.value.trim();
                    tab.savedAltVer = tab.model.getAlternativeVersionId();
                    close();
                    this.#logMessage('ok', `Saved stored query "${qName}".`);
                    this.#renderAll();
                } catch (err) {
                    this.#logMessage('err', `Save query failed: ${err.message}`);
                    close();
                    this.#activeBottomTab = 'messages';
                    this.#renderResultsPane();
                }
            });
        }

        promptSaveQuery() {
            const tab = this.activeTab;
            if (!tab) return;
            const modal = document.getElementById('ss-save-modal');
            const nameIn = /** @type {HTMLInputElement} */ (document.getElementById('ss-q-name'));
            const catIn = /** @type {HTMLInputElement} */ (document.getElementById('ss-q-cat'));
            const descIn = /** @type {HTMLInputElement} */ (document.getElementById('ss-q-desc'));

            nameIn.value = tab.savedName || (tab.title.startsWith('Query ') ? '' : tab.title);
            catIn.value = tab.category || 'General';
            descIn.value = tab.description || '';
            modal.hidden = false;
            setTimeout(() => nameIn.focus(), 20);
        }

        async deleteSavedQuery(qName, ev) {
            if (ev) ev.stopPropagation();
            try {
                const res = await bridge.call('axlite.deleteQuery', { name: qName });
                this.#savedQueries = res.savedQueries || [];
                this.#logMessage('ok', `Deleted stored query "${qName}".`);
                this.#renderTree();
                this.#renderResultsPane();
            } catch (err) {
                this.#logMessage('err', `Delete query failed: ${err.message}`);
                this.#activeBottomTab = 'messages';
                this.#renderResultsPane();
            }
        }

        toggleTheme() {
            this.#darkMode = !this.#darkMode;
            document.body.classList.toggle('theme-light', !this.#darkMode);
            this.#monaco.editor.setTheme(this.#darkMode ? 'vs-dark' : 'vs');
            this.#renderTopbar();
        }

        #logMessage(kind, text) {
            this.#messages.unshift({
                kind,
                time: new Date().toLocaleTimeString(),
                text
            });
            if (this.#messages.length > 100) this.#messages.length = 100;
        }

        // =====================================================================
        // Rendering
        // =====================================================================

        #renderAll() {
            this.#renderTopbar();
            this.#renderTree();
            this.#renderQueryTabs();
            this.#renderResultsPane();
            this.#renderStatusBar();
        }

        #renderTopbar() {
            const limitSelect = el('select', {
                class: 'ss-select',
                title: 'Maximum SELECT rows fetched into grid',
                on: {
                    change: (e) => { this.#rowLimit = Number(e.target.value); }
                }
            },
                [200, 1000, 5000, 0].map((val) => {
                    const opt = el('option', { value: val }, val === 0 ? 'Limit: All' : `Limit: ${val}`);
                    if (val === this.#rowLimit) /** @type {HTMLOptionElement} */ (opt).selected = true;
                    return opt;
                })
            );

            const maintSelect = el('select', {
                class: 'ss-select',
                title: 'SQLite Database Maintenance Tools',
                on: {
                    change: async (e) => {
                        const action = e.target.value;
                        e.target.value = '';
                        if (action) await this.runDbMaintenance(action);
                    }
                }
            },
                el('option', { value: '' }, 'DB Tools...'),
                el('option', { value: 'integrity_check' }, 'PRAGMA integrity_check'),
                el('option', { value: 'vacuum' }, 'VACUUM Database'),
                el('option', { value: 'backup' }, 'Online Backup to File...')
            );

            this.#topbarEl.replaceChildren(
                el('div', { class: 'ss-brand' },
                    el('span', { class: 'ss-brand-dot' }),
                    'AxLite'
                ),
                el('div', { class: 'ss-db-pill', title: this.#dbInfo.dbPath || ':memory:' },
                    this.#dbInfo.dbName || '(No DB)'
                ),
                el('button', { type: 'button', class: 'ss-btn', title: 'Open an existing SQLite database file', on: { click: () => this.openDatabaseDialog('open') } }, 'Open DB...'),
                el('button', { type: 'button', class: 'ss-btn', title: 'Create a new SQLite database file', on: { click: () => this.openDatabaseDialog('create') } }, 'New DB...'),
                el('button', { type: 'button', class: 'ss-btn', title: 'Open an isolated in-memory SQLite database', on: { click: () => this.openInMemoryDb() } }, ':memory:'),
                el('span', { class: 'ss-spacer' }),
                el('button', {
                    type: 'button',
                    class: 'ss-btn ss-btn-primary',
                    disabled: this.#busy,
                    title: 'Execute highlighted selection or active tab (Ctrl+Enter / F5)',
                    on: { click: () => this.runActiveSql() }
                }, this.#busy ? 'Running...' : '▶ Run SQL'),
                el('button', {
                    type: 'button',
                    class: 'ss-btn',
                    disabled: this.#busy,
                    title: 'Run EXPLAIN QUERY PLAN on current SQL',
                    on: { click: () => this.explainActiveSql() }
                }, 'Explain Plan'),
                el('button', {
                    type: 'button',
                    class: 'ss-btn',
                    title: 'Save current SQL into Stored Queries (Ctrl+S)',
                    on: { click: () => this.promptSaveQuery() }
                }, 'Save Query'),
                limitSelect,
                maintSelect,
                el('button', { type: 'button', class: 'ss-btn', title: 'Refresh tables, views, and stored queries', on: { click: () => this.refreshSchema() } }, '↻'),
                el('button', { type: 'button', class: 'ss-btn', title: 'Toggle Dark / Light Editor Theme', on: { click: () => this.toggleTheme() } }, this.#darkMode ? '☀' : '☾')
            );
        }

        #matchesFilter(name, childrenNames = []) {
            if (!this.#treeFilter) return true;
            if (String(name).toLowerCase().includes(this.#treeFilter)) return true;
            return childrenNames.some((c) => String(c).toLowerCase().includes(this.#treeFilter));
        }

        #renderTree() {
            const nodes = [];

            // 1. Stored Queries
            const visibleQueries = (this.#savedQueries || []).filter((q) =>
                this.#matchesFilter(q.name, [q.category || '', q.description || '', q.sql || ''])
            );
            nodes.push(el('div', { class: 'ss-sec-hdr' },
                el('span', {}, 'Stored Queries'),
                el('span', { class: 'ss-sec-count' }, String(visibleQueries.length))
            ));
            if (visibleQueries.length === 0) {
                nodes.push(el('div', { class: 'ss-node-row', style: { cursor: 'default', color: 'var(--text-muted)' } },
                    'No stored queries (Ctrl+S to save)'
                ));
            } else {
                for (const q of visibleQueries) {
                    nodes.push(el('div', {
                        class: 'ss-node-row',
                        title: `${q.category ? '[' + q.category + '] ' : ''}${q.description || q.sql}`,
                        on: {
                            click: () => this.openSqlInNewTab(q.name, q.sql, {
                                savedName: q.name,
                                category: q.category,
                                description: q.description
                            })
                        }
                    },
                        el('span', { class: 'ss-node-caret' }, '★'),
                        el('span', { class: 'ss-node-name' }, q.name),
                        el('span', { class: 'ss-node-actions' },
                            el('button', {
                                type: 'button',
                                class: 'ss-mini-btn',
                                title: 'Open & Execute Stored Query',
                                on: {
                                    click: (ev) => {
                                        ev.stopPropagation();
                                        this.openSqlInNewTab(q.name, q.sql, {
                                            savedName: q.name,
                                            category: q.category,
                                            description: q.description
                                        });
                                        this.runActiveSql(q.sql);
                                    }
                                }
                            }, 'Run'),
                            el('button', {
                                type: 'button',
                                class: 'ss-mini-btn',
                                title: 'Delete Stored Query',
                                on: { click: (ev) => this.deleteSavedQuery(q.name, ev) }
                            }, '×')
                        )
                    ));
                }
            }

            // 2. Tables & Columns
            const visibleTables = (this.#schema.tables || []).filter((t) =>
                this.#matchesFilter(t.name, (t.columns || []).map((c) => c.name))
            );
            nodes.push(el('div', { class: 'ss-sec-hdr', style: { marginTop: '8px' } },
                el('span', {}, 'Tables'),
                el('span', { class: 'ss-sec-count' }, String(visibleTables.length))
            ));

            for (const tbl of visibleTables) {
                const key = `tbl:${tbl.name}`;
                const expanded = this.#expandedNodes.has(key) || Boolean(this.#treeFilter);
                const fkCols = new Set((tbl.foreignKeys || []).map((fk) => String(fk.from).toLowerCase()));

                nodes.push(el('div', {
                    class: 'ss-node-row',
                    on: {
                        click: () => {
                            if (this.#expandedNodes.has(key)) this.#expandedNodes.delete(key);
                            else this.#expandedNodes.add(key);
                            this.#renderTree();
                        }
                    }
                },
                    el('span', { class: 'ss-node-caret' }, expanded ? '▾' : '▸'),
                    el('span', { class: 'ss-node-name', title: tbl.name }, tbl.name),
                    tbl.strict ? el('span', { class: 'ss-badge ss-badge-nn' }, 'STRICT') : null,
                    el('span', { class: 'ss-node-meta' }, `${tbl.rowCount ?? 0}`),
                    el('span', { class: 'ss-node-actions' },
                        el('button', {
                            type: 'button',
                            class: 'ss-mini-btn',
                            title: 'Browse Top 200 Rows',
                            on: {
                                click: (ev) => {
                                    ev.stopPropagation();
                                    const sql = `SELECT * FROM ${quoteIdent(tbl.name)} LIMIT 200;`;
                                    this.openSqlInNewTab(tbl.name, sql);
                                    this.runActiveSql(sql);
                                }
                            }
                        }, 'Rows'),
                        el('button', {
                            type: 'button',
                            class: 'ss-mini-btn',
                            title: 'Generate INSERT template',
                            on: {
                                click: (ev) => {
                                    ev.stopPropagation();
                                    const editableCols = (tbl.columns || []).filter((c) => !c.pk);
                                    const colsList = (editableCols.length ? editableCols : (tbl.columns || [])).map((c) => quoteIdent(c.name));
                                    const valsList = colsList.map((c) => `/* ${c} */ NULL`);
                                    const sql = `INSERT INTO ${quoteIdent(tbl.name)} (\n    ${colsList.join(',\n    ')}\n) VALUES (\n    ${valsList.join(',\n    ')}\n);\n`;
                                    this.openSqlInNewTab(`Insert_${tbl.name}`, sql);
                                }
                            }
                        }, 'Ins'),
                        el('button', {
                            type: 'button',
                            class: 'ss-mini-btn',
                            title: 'Inspect CREATE TABLE DDL',
                            on: {
                                click: (ev) => {
                                    ev.stopPropagation();
                                    const idxDdl = (tbl.indexes || []).filter((ix) => ix.sql).map((ix) => `${ix.sql};`).join('\n');
                                    const ddl = `-- Table DDL: ${tbl.name}\n${tbl.sql};\n${idxDdl ? '\n-- Indexes:\n' + idxDdl + '\n' : ''}`;
                                    this.openSqlInNewTab(`DDL_${tbl.name}`, ddl);
                                }
                            }
                        }, 'DDL')
                    )
                ));

                if (expanded) {
                    const colItems = (tbl.columns || []).map((col) => {
                        const isFk = fkCols.has(String(col.name).toLowerCase());
                        return el('div', {
                            class: 'ss-col-item',
                            title: `Click to insert ${col.name} into SQL editor`,
                            on: { click: () => this.#insertAtCursor(quoteIdent(col.name)) }
                        },
                            el('span', {}, col.name),
                            col.pk ? el('span', { class: 'ss-badge ss-badge-pk' }, 'PK') : null,
                            isFk ? el('span', { class: 'ss-badge ss-badge-fk' }, 'FK') : null,
                            col.notnull && !col.pk ? el('span', { class: 'ss-badge ss-badge-nn' }, 'NN') : null,
                            el('span', { class: 'ss-col-type' }, col.type || 'ANY')
                        );
                    });

                    for (const ix of tbl.indexes || []) {
                        colItems.push(el('div', {
                            class: 'ss-col-item',
                            style: { color: 'var(--text-secondary)' },
                            title: `Index: ${ix.name} (${ix.columns || ''})`
                        },
                            el('span', {}, `⌗ ${ix.name}`),
                            ix.unique ? el('span', { class: 'ss-badge ss-badge-nn' }, 'UQ') : null,
                            el('span', { class: 'ss-col-type' }, ix.columns || '')
                        ));
                    }

                    nodes.push(el('div', { class: 'ss-col-list' }, colItems));
                }
            }

            // 3. Views
            const visibleViews = (this.#schema.views || []).filter((v) =>
                this.#matchesFilter(v.name, (v.columns || []).map((c) => c.name))
            );
            nodes.push(el('div', { class: 'ss-sec-hdr', style: { marginTop: '8px' } },
                el('span', {}, 'Views'),
                el('span', { class: 'ss-sec-count' }, String(visibleViews.length))
            ));
            for (const vw of visibleViews) {
                const key = `vw:${vw.name}`;
                const expanded = this.#expandedNodes.has(key);
                nodes.push(el('div', {
                    class: 'ss-node-row',
                    on: {
                        click: () => {
                            if (this.#expandedNodes.has(key)) this.#expandedNodes.delete(key);
                            else this.#expandedNodes.add(key);
                            this.#renderTree();
                        }
                    }
                },
                    el('span', { class: 'ss-node-caret' }, expanded ? '▾' : '▸'),
                    el('span', { class: 'ss-node-name' }, vw.name),
                    el('span', { class: 'ss-node-actions' },
                        el('button', {
                            type: 'button',
                            class: 'ss-mini-btn',
                            on: {
                                click: (ev) => {
                                    ev.stopPropagation();
                                    const sql = `SELECT * FROM ${quoteIdent(vw.name)} LIMIT 200;`;
                                    this.openSqlInNewTab(vw.name, sql);
                                    this.runActiveSql(sql);
                                }
                            }
                        }, 'Rows'),
                        el('button', {
                            type: 'button',
                            class: 'ss-mini-btn',
                            on: {
                                click: (ev) => {
                                    ev.stopPropagation();
                                    this.openSqlInNewTab(`View_${vw.name}`, `${vw.sql};\n`);
                                }
                            }
                        }, 'DDL')
                    )
                ));
                if (expanded) {
                    nodes.push(el('div', { class: 'ss-col-list' },
                        (vw.columns || []).map((col) => el('div', {
                            class: 'ss-col-item',
                            on: { click: () => this.#insertAtCursor(quoteIdent(col.name)) }
                        },
                            el('span', {}, col.name),
                            el('span', { class: 'ss-col-type' }, col.type || '')
                        ))
                    ));
                }
            }

            // 4. Triggers (if any)
            const visibleTriggers = (this.#schema.triggers || []).filter((tr) => this.#matchesFilter(tr.name, [tr.tbl_name]));
            if (visibleTriggers.length > 0) {
                nodes.push(el('div', { class: 'ss-sec-hdr', style: { marginTop: '8px' } },
                    el('span', {}, 'Triggers'),
                    el('span', { class: 'ss-sec-count' }, String(visibleTriggers.length))
                ));
                for (const tr of visibleTriggers) {
                    nodes.push(el('div', {
                        class: 'ss-node-row',
                        on: { click: () => this.openSqlInNewTab(`Trigger_${tr.name}`, `${tr.sql};\n`) }
                    },
                        el('span', { class: 'ss-node-caret' }, '⚡'),
                        el('span', { class: 'ss-node-name' }, tr.name),
                        el('span', { class: 'ss-node-meta' }, tr.tbl_name)
                    ));
                }
            }

            this.#treeEl.replaceChildren(...nodes);
        }

        #renderQueryTabs() {
            const tabButtons = this.#tabs.map((t) => {
                const dirty = t.model.getAlternativeVersionId() !== t.savedAltVer;
                return el('div', {
                    class: `ss-qtab${t.id === this.#activeTabId ? ' active' : ''}`,
                    on: { click: () => this.#switchTab(t.id) }
                },
                    el('span', {}, `${t.title}${dirty ? ' ●' : ''}`),
                    el('span', {
                        class: 'ss-qtab-close',
                        title: 'Close tab',
                        on: { click: (ev) => this.#closeTab(t.id, ev) }
                    }, '×')
                );
            });

            tabButtons.push(el('button', {
                type: 'button',
                class: 'ss-btn',
                style: { height: '24px', marginLeft: '4px', padding: '0 8px' },
                title: 'Open new SQL editor tab',
                on: { click: () => this.#createTab(`Query ${this.#tabSeq + 1}`, '') }
            }, '+'));

            this.#queryTabsEl.replaceChildren(...tabButtons);
        }

        #getSortedRows() {
            if (!this.#lastResult || !this.#lastResult.rows) return [];
            const rows = [...this.#lastResult.rows];
            const idx = this.#sortColIdx;
            if (idx < 0) return rows;
            const dir = this.#sortAsc ? 1 : -1;
            rows.sort((a, b) => {
                const va = a[idx];
                const vb = b[idx];
                if (va === vb) return 0;
                if (va === null || va === undefined) return 1;
                if (vb === null || vb === undefined) return -1;
                if (typeof va === 'number' && typeof vb === 'number') return (va - vb) * dir;
                return String(va).localeCompare(String(vb), undefined, { numeric: true }) * dir;
            });
            return rows;
        }

        #renderResultsPane() {
            const hasRows = Boolean(this.#lastResult && this.#lastResult.columns && this.#lastResult.columns.length);
            const resSummary = hasRows
                ? `${this.#lastResult.rowCount} row(s)${this.#lastResult.truncated ? ' (truncated at limit)' : ''} · ${Number(this.#lastResult.totalElapsedMs || 0).toFixed(2)} ms`
                : '';

            setChildren(this.#resBarEl,
                el('button', {
                    type: 'button',
                    class: `ss-rtab${this.#activeBottomTab === 'results' ? ' active' : ''}`,
                    on: { click: () => { this.#activeBottomTab = 'results'; this.#renderResultsPane(); } }
                }, `Results${hasRows ? ` (${this.#lastResult.rowCount})` : ''}`),
                el('button', {
                    type: 'button',
                    class: `ss-rtab${this.#activeBottomTab === 'messages' ? ' active' : ''}`,
                    on: { click: () => { this.#activeBottomTab = 'messages'; this.#renderResultsPane(); } }
                }, `Messages (${this.#messages.length})`),
                el('button', {
                    type: 'button',
                    class: `ss-rtab${this.#activeBottomTab === 'plan' ? ' active' : ''}`,
                    on: { click: () => { this.#activeBottomTab = 'plan'; this.#renderResultsPane(); } }
                }, 'Query Plan'),
                el('button', {
                    type: 'button',
                    class: `ss-rtab${this.#activeBottomTab === 'history' ? ' active' : ''}`,
                    on: { click: () => { this.#activeBottomTab = 'history'; this.#renderResultsPane(); } }
                }, `History (${this.#history.length})`),
                el('span', { class: 'ss-spacer' }),
                resSummary ? el('span', { class: 'ss-res-meta' }, resSummary) : null,
                hasRows ? el('button', { type: 'button', class: 'ss-btn', on: { click: () => this.exportResultsCsv() } }, 'Export CSV...') : null,
                hasRows ? el('button', { type: 'button', class: 'ss-btn', on: { click: () => this.copyResultsJson() } }, 'Copy JSON') : null
            );

            if (this.#activeBottomTab === 'results') {
                if (!hasRows) {
                    this.#resBodyEl.replaceChildren(el('div', { class: 'ss-empty' },
                        'Run a SELECT, PRAGMA, EXPLAIN, or RETURNING query (Ctrl+Enter / F5) to view results here.'
                    ));
                    return;
                }

                const cols = this.#lastResult.columns;
                const rows = this.#getSortedRows();

                const thead = el('thead', {}, el('tr', {},
                    el('th', { class: 'ss-row-num' }, '#'),
                    cols.map((col, idx) => {
                        const arrow = this.#sortColIdx === idx ? (this.#sortAsc ? ' ▲' : ' ▼') : '';
                        return el('th', {
                            title: 'Click to sort column',
                            on: {
                                click: () => {
                                    if (this.#sortColIdx === idx) this.#sortAsc = !this.#sortAsc;
                                    else { this.#sortColIdx = idx; this.#sortAsc = true; }
                                    this.#renderResultsPane();
                                }
                            }
                        },
                            `${col.name}${arrow}`,
                            col.declType ? el('span', { class: 'ss-th-type' }, col.declType) : null
                        );
                    })
                ));

                const tbody = el('tbody', {}, rows.map((r, rIdx) => el('tr', {},
                    el('td', { class: 'ss-row-num' }, String(rIdx + 1)),
                    cols.map((_c, cIdx) => {
                        const val = r[cIdx];
                        if (val === null || val === undefined) {
                            return el('td', {}, el('span', { class: 'ss-null' }, 'NULL'));
                        }
                        if (typeof val === 'number') {
                            return el('td', { class: 'ss-num', title: String(val) }, String(val));
                        }
                        const text = String(val);
                        return el('td', { title: text.length > 60 ? text : '' }, text);
                    })
                )));

                this.#resBodyEl.replaceChildren(el('table', { class: 'ss-grid' }, thead, tbody));
            } else if (this.#activeBottomTab === 'messages') {
                if (!this.#messages.length) {
                    this.#resBodyEl.replaceChildren(el('div', { class: 'ss-empty' }, 'No execution messages yet.'));
                    return;
                }
                this.#resBodyEl.replaceChildren(el('div', { class: 'ss-msg-list' },
                    this.#messages.map((m) => el('div', {
                        class: `ss-msg-card ${m.kind === 'err' ? 'err' : 'ok'}`,
                        style: { whiteSpace: 'pre-wrap' }
                    }, `[${m.time}] ${m.text}`))
                ));
            } else if (this.#activeBottomTab === 'plan') {
                if (!this.#planNodes.length) {
                    this.#resBodyEl.replaceChildren(el('div', { class: 'ss-empty' },
                        'Click "Explain Plan" in the toolbar to inspect SQLite\'s EXPLAIN QUERY PLAN tree.'
                    ));
                    return;
                }
                this.#resBodyEl.replaceChildren(el('div', { class: 'ss-msg-list' },
                    this.#planNodes.map((n) => el('div', { class: 'ss-msg-card' },
                        `[id=${n.id}, parent=${n.parent}]  ${n.detail}`
                    ))
                ));
            } else if (this.#activeBottomTab === 'history') {
                if (!this.#history.length) {
                    this.#resBodyEl.replaceChildren(el('div', { class: 'ss-empty' }, 'No SQL statements executed in this session yet.'));
                    return;
                }
                this.#resBodyEl.replaceChildren(el('div', { class: 'ss-msg-list' },
                    this.#history.map((hItem, idx) => el('div', {
                        class: `ss-msg-card ${hItem.ok ? 'ok' : 'err'}`,
                        style: { cursor: 'pointer' },
                        title: 'Click to open this SQL statement in a new editor tab',
                        on: { click: () => this.openSqlInNewTab(`History_${idx + 1}`, hItem.sql) }
                    },
                        `[${hItem.time}] ${hItem.summary}${hItem.elapsedMs ? ` (${Number(hItem.elapsedMs).toFixed(2)} ms)` : ''}\n${hItem.sql}`
                    ))
                ));
            }
        }

        #renderStatusBar(flashNote = '') {
            const pos = this.#editor ? this.#editor.getPosition() : null;
            setChildren(this.#statusEl,
                el('span', { class: 'ss-status-ok' },
                    this.#dbInfo.engineVersion ? `SQLite ${this.#dbInfo.engineVersion}` : 'Connecting...'
                ),
                this.#dbInfo.journalMode ? el('span', {}, `journal_mode: ${this.#dbInfo.journalMode.toUpperCase()}`) : null,
                el('span', {}, `FK: ${this.#dbInfo.foreignKeys ? 'ON' : 'OFF'}`),
                el('span', {}, `Size: ${formatBytes(this.#dbInfo.fileSizeBytes)}`),
                this.#dbInfo.inTransaction ? el('span', { class: 'ss-status-tx' }, '● TRANSACTION OPEN') : null,
                el('span', { class: 'ss-spacer' }),
                flashNote ? el('span', { class: 'ss-status-ok' }, flashNote) : null,
                pos ? el('span', {}, `Ln ${pos.lineNumber}, Col ${pos.column}`) : null,
                el('span', {}, 'Ctrl+Enter / F5: Run · Ctrl+S: Save Stored Query')
            );
        }
    }

    async function boot() {
        const root = document.getElementById('ss-app');
        try {
            const monaco = await loadMonaco();
            AxLite.app = new AxLiteApp(root, monaco);
        } catch (err) {
            root.replaceChildren(el('div', { class: 'ss-msg-list' },
                el('div', { class: 'ss-msg-card err' }, `Failed to initialize AxLite: ${err.message || err}`)
            ));
        }
    }

    boot();
})(window.AxLite = window.AxLite || {});
