window.excelInterop = {
    isReady: false,
    host: null,
    skipped: false,

    isOfficeEnvironment() {
        if (window.location.href.includes("_host_Info")) {
            return true;
        }

        try {
            if (window.sessionStorage.getItem("hostInfoValue")) {
                return true;
            }
        } catch {
            // sessionStorage can throw in a sandboxed frame
        }

        return false;
    },

    start() {
        this._nativePushState = window.history.pushState.bind(window.history);
        this._nativeReplaceState = window.history.replaceState.bind(window.history);

        if (!this.isOfficeEnvironment()) {
            this.skipped = true;
            this.isReady = false;
            this.host = null;
            return;
        }

        window.Office = window.Office || {};
        if (typeof window.Office.initialize !== "function") {
            window.Office.initialize = function () { };
        }

        const officeBase = "https://appsforoffice.microsoft.com/lib/1/hosted/";
        const afterOfficeJs = () => {
            try {
                Office.initialize = Office.initialize || function () { };
            } catch {
                // Office may still be constructing its namespace.
            }

            Office.onReady((info) => {
                this.host = info.host ? String(info.host) : null;
                this.isReady = true;
                this._startBlazorIfPending();
            });
        };

        const loadOfficeJs = () => this._loadScript(officeBase + "office.js")
            .then(afterOfficeJs)
            .catch(() => {
                this.isReady = false;
                this.host = null;
                this._startBlazorIfPending();
            });

        this._loadScript(officeBase + "de-de/office_strings.js")
            .catch(() => {})
            .then(loadOfficeJs);
    },

    startBlazor() {
        const begin = () => {
            if (this._blazorStarted || typeof Blazor === "undefined") {
                return;
            }

            this._blazorStarted = true;
            Promise.resolve(Blazor.start({ applicationCulture: "de-DE" })).catch((error) => {
                console.error("SPMT4all: Blazor.start failed", error);
            });
        };

        if (this.skipped || this.isReady) {
            begin();
            return;
        }

        this._pendingBlazorStart = begin;
        window.setTimeout(() => this._startBlazorIfPending(), 15000);
    },

    _startBlazorIfPending() {
        if (typeof this._pendingBlazorStart === "function") {
            const begin = this._pendingBlazorStart;
            this._pendingBlazorStart = null;
            begin();
        }
    },

    _loadScript(src) {
        return new Promise((resolve, reject) => {
            const script = document.createElement("script");
            script.src = src;
            script.async = false;
            script.onload = () => resolve();
            script.onerror = () => reject(new Error("Failed to load " + src));
            document.head.appendChild(script);
        });
    },

    async waitForHost(timeoutMs = 8000) {
        if (this.skipped) {
            return null;
        }

        const start = Date.now();
        while (!this.isReady && Date.now() - start < timeoutMs) {
            await new Promise((resolve) => setTimeout(resolve, 50));
        }

        return this.host;
    },

    getStatus() {
        return {
            ok: this.isReady && this.host === "Excel",
            host: this.host,
            skipped: this.skipped
        };
    },

    /**
     * Updates the task-pane URL for in-pane Blazor routes without leaving the GitHub Pages base path.
     */
    navigate(path) {
        const url = new URL(window.location.href);
        const baseHref = document.querySelector("base")?.getAttribute("href") || "/";
        const baseUrl = new URL(baseHref, url.origin);
        const relative = !path || path === "/" ? "./" : String(path).replace(/^\//, "");
        const nextUrl = new URL(relative, baseUrl);
        nextUrl.search = url.search;
        nextUrl.hash = "";
        const next = `${nextUrl.pathname}${nextUrl.search}`;
        const push = this._nativePushState || window.history.pushState.bind(window.history);
        push(null, "", next);
        window.dispatchEvent(new PopStateEvent("popstate"));
    },

    /**
     * Writes a string value to a cell address on the active worksheet (e.g. "A1").
     */
    async writeCell(address, value) {
        await Excel.run(async (context) => {
            const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
            range.values = [[value == null ? "" : String(value)]];
            await context.sync();
        });
    },

    async readCell(address) {
        return await Excel.run(async (context) => {
            const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
            range.load("values");
            await context.sync();
            const cell = range.values && range.values[0] ? range.values[0][0] : null;
            return cell == null ? null : String(cell);
        });
    },

    /**
     * Reads the used range of a named sheet as string cells (for settings storage).
     * Returns null when the sheet does not exist.
     */
    async readUsedRange(sheetName) {
        return await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            sheet.load("isNullObject");
            await context.sync();
            if (sheet.isNullObject) {
                return null;
            }

            const range = sheet.getUsedRangeOrNullObject();
            range.load(["values", "rowCount", "columnCount"]);
            await context.sync();
            if (range.isNullObject) {
                return { values: [], rowCount: 0, columnCount: 0 };
            }

            return {
                values: range.values.map(row => row.map(cell => cell == null ? null : String(cell))),
                rowCount: range.rowCount,
                columnCount: range.columnCount
            };
        });
    },

    /**
     * Creates or replaces a named sheet with a full 2D values grid (settings storage).
     * Sheets are very-hidden unless showSheet is true (Debug).
     */
    async writeSheet(sheetName, values, showSheet) {
        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            sheet.load("isNullObject");
            await context.sync();

            if (sheet.isNullObject) {
                sheet = context.workbook.worksheets.add(sheetName);
            }

            sheet.visibility = showSheet
                ? Excel.SheetVisibility.visible
                : Excel.SheetVisibility.veryHidden;

            const used = sheet.getUsedRangeOrNullObject();
            used.load("isNullObject");
            await context.sync();
            if (!used.isNullObject) {
                used.clear();
            }

            const rowCount = values.length;
            const columnCount = rowCount === 0 ? 1 : values[0].length;
            const target = sheet.getRangeByIndexes(0, 0, Math.max(rowCount, 1), Math.max(columnCount, 1));
            target.numberFormat = "@";
            if (rowCount > 0) {
                target.values = values;
                sheet.getRangeByIndexes(0, 0, 1, columnCount).format.font.bold = true;
            }

            await context.sync();
        });
    },

    /**
     * Sets visibility of an existing named sheet (visible vs veryHidden).
     */
    async setSheetVisibility(sheetName, showSheet) {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            sheet.load("isNullObject");
            await context.sync();
            if (sheet.isNullObject) {
                return;
            }

            sheet.visibility = showSheet
                ? Excel.SheetVisibility.visible
                : Excel.SheetVisibility.veryHidden;
            await context.sync();
        });
    }
};

window.excelInterop.start();
