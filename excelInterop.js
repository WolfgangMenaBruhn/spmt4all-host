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

        this._loadScript(officeBase + "en-us/office_strings.js")
            .catch(() => {})
            .then(loadOfficeJs);
    },

    startBlazor() {
        const begin = () => {
            if (this._blazorStarted || typeof Blazor === "undefined") {
                return;
            }

            this._blazorStarted = true;
            Promise.resolve(Blazor.start({ applicationCulture: "en" })).catch((error) => {
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
    }
};

window.excelInterop.start();
