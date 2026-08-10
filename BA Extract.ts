// =======================================================
// OFFICE SCRIPT: GL Code Processor (Column F + AH Output)
// =======================================================

function main(workbook: ExcelScript.Workbook) {
    processGLCodes(workbook);
}

function loadProhibitedGLs(workbook: ExcelScript.Workbook): Map<string, string> {
    const sheet = workbook.getWorksheet("Prohibited GLs");
    if (!sheet) return new Map();
    const used = sheet.getUsedRange();
    if (!used || used.getRowCount() < 2) return new Map();
    const values = sheet.getRangeByIndexes(1, 0, used.getRowCount() - 1, 2).getValues();
    const map = new Map<string, string>();
    for (const [code, desc] of values) {
        if (code) map.set(String(code).trim(), String(desc).trim());
    }
    return map;
}

function loadMasterGLMap(workbook: ExcelScript.Workbook): Map<string, string> {
    const sheet = workbook.getWorksheet("Mapping - EN");
    if (!sheet) return new Map();
    const used = sheet.getUsedRange();
    if (!used || used.getRowCount() < 2) return new Map();
    const values = sheet.getRangeByIndexes(1, 0, used.getRowCount() - 1, 2).getValues();
    const map = new Map<string, string>();
    for (const [code, desc] of values) {
        if (code) map.set(String(code).trim(), desc ? String(desc).trim() : "(description not found)");
    }
    return map;
}

function processGLCodes(workbook: ExcelScript.Workbook): void {
    const ws = workbook.getWorksheet("2026-27 Tracker (Testing)");
    if (!ws) return;

    const prohibitedMap = loadProhibitedGLs(workbook);
    const masterGLMap = loadMasterGLMap(workbook);

    const used = ws.getUsedRange();
    if (!used) return;

    const rowCount = used.getRowCount() - 1;
    if (rowCount < 1) return;

    const colE = ws.getRangeByIndexes(1, 4, rowCount, 1).getValues();
    const colN = ws.getRangeByIndexes(1, 13, rowCount, 1).getValues();
    const colO = ws.getRangeByIndexes(1, 14, rowCount, 1).getValues();
    const colQ = ws.getRangeByIndexes(1, 16, rowCount, 1).getValues();

    const colF_existing = ws.getRangeByIndexes(1, 5, rowCount, 1).getValues();
    const colAH_existing = ws.getRangeByIndexes(1, 33, rowCount, 1).getValues();

    const outputF: string[][] = colF_existing.map(row => [String(row[0] ?? "")]);
    const outputAH: string[][] = colAH_existing.map(row => [String(row[0] ?? "")]);

    for (let i = 0; i < rowCount; i++) {

        const rawText = String(colO[i][0]).trim();
        if (!rawText) continue;

        const lines = rawText.replace(/\r/g, "").split("\n");

        const gls: string[] = [];
        const seen = new Set<string>();
        const duplicates = new Set<string>();

        interface GLEntry { code: string; descFromO: string; }
        const glEntries: GLEntry[] = [];

        let currentCode = "";
        let currentDesc = "";

        for (const line of lines) {
            const trimmed = line.trim();
            if (trimmed.startsWith("GL Account Code:")) {
                currentCode = trimmed.substring(16).trim();
            } else if (trimmed.startsWith("GL Account Description:")) {
                currentDesc = trimmed.substring(23).trim();
            } else if (trimmed === "---") {
                if (currentCode) {
                    glEntries.push({ code: currentCode, descFromO: currentDesc });
                    currentCode = "";
                    currentDesc = "";
                }
            }
        }
        if (currentCode) {
            glEntries.push({ code: currentCode, descFromO: currentDesc });
        }

        for (const entry of glEntries) {
            const code = entry.code;
            gls.push(code);
            seen.has(code) ? duplicates.add(code) : seen.add(code);
        }

        // ---------- CCG flag: cost centre range 600000–750000 ----------
        const costCentreRaw = String(colN[i][0]).trim();
        const costCentreNum = Number(costCentreRaw);
        const isCCG =
            !isNaN(costCentreNum) &&
            costCentreNum >= 600000 &&
            costCentreNum <= 750000;

        if (isCCG) {
            // Column AH: write CCG if blank
            if (String(colAH_existing[i][0]).trim() === "") {
                outputAH[i][0] = "CCG";
            }

            // Column F: write CCM name (from Column Q email) if blank
            if (String(colF_existing[i][0]).trim() === "") {
                const colEValue = String(colE[i][0]);
                const pendingApprovalPattern = /PENDING\s+APPROVAL\s*:\s*Internal\s+Controls['\u2019]?\s*Approval/i;
                if (pendingApprovalPattern.test(colEValue)) {
                    const email = String(colQ[i][0]).trim();
                    let ccmName = "";
                    if (email.includes("@") && email.includes(".")) {
                        const local = email.split("@")[0];
                        const [f, l] = local.split(".");
                        if (f && l) {
                            ccmName = `${toProperCase(f)} ${toProperCase(l)}`;
                        }
                    }
                    outputF[i][0] = ccmName
                        ? `${ccmName}`
                        : `Note: CCG`;
                }
            }
            // Skip all GL code processing for this row
            continue;
        }

        // ---------- Column AH: write only if blank ----------
        if (String(colAH_existing[i][0]).trim() === "" && glEntries.length > 0) {
            const ahLines: string[] = [];
            ahLines.push("[GL account code from Mapping - EN] [GL account description from Mapping - EN]\n[GL account code from column O] [GL account description from column O]");

            let allPairsMatch = true;

            for (const entry of glEntries) {
                const mappedDesc = masterGLMap.get(entry.code) ?? "(not in Mapping - EN)";
                const mappingLine = `${entry.code} ${mappedDesc}`;
                const colOLine = `${entry.code} ${entry.descFromO || "(no description in application)"}`;
                if (mappingLine !== colOLine) allPairsMatch = false;
                ahLines.push(`${mappingLine}\n${colOLine}`);
            }

            outputAH[i][0] = ahLines.join("\n\n");

            if (allPairsMatch) {
                outputAH[i][0] = "GL Account Codes match GL Account Descriptions";
            }
        }

        // ---------- Column F: skip if already has content ----------
        if (String(colF_existing[i][0]).trim() !== "") continue;

        // ---------- Column F: skip if Column E does not contain the required approval status ----------
        const colEValue = String(colE[i][0]);
        const pendingApprovalPattern = /PENDING\s+APPROVAL\s*:\s*Internal\s+Controls['\u2019]?\s*Approval/i;
        if (!pendingApprovalPattern.test(colEValue)) continue;
        if (gls.length === 0) continue;

        // ---------- CCM name from Column Q email ----------
        let ccmName = "";
        const email = String(colQ[i][0]).trim();
        if (email.includes("@") && email.includes(".")) {
            const local = email.split("@")[0];
            const [f, l] = local.split(".");
            if (f && l) {
                ccmName = `${toProperCase(f)} ${toProperCase(l)}`;
            }
        }

        // ---------- Classify ----------
        const prohibited: string[] = [];
        const typos: string[] = [];

        for (const gl of gls) {
            if (!masterGLMap.has(gl)) {
                typos.push(gl);
            } else if (prohibitedMap.has(gl)) {
                prohibited.push(`${gl} (${prohibitedMap.get(gl)}) –`);
            }
        }

        // ---------- Header ----------
        const parts: string[] = [];
        parts.push(`${gls.length} GL${gls.length > 1 ? "s" : ""} on application`);

        if (prohibited.length > 0) {
            parts.push(`${prohibited.length} prohibited`);
        } else {
            parts.push(`none prohibited`);
        }

        if (typos.length > 0) {
            parts.push(`${typos.length} typo${typos.length > 1 ? "s" : ""}`);
        }

        if (duplicates.size > 0) {
            const d = duplicates.size;
            parts.push(`${d} duplicate${d > 1 ? "s" : ""}`);
        }

        // ---------- Body ----------
        const body: string[] = [];

        body.push(...prohibited);

        typos.forEach(gl => {
            const entry = glEntries.find(e => e.code === gl);
            const descFromO = entry?.descFromO || "";
            const label = descFromO ? `${gl} (for "${descFromO}")` : gl;
            body.push(`${label} – error, GL code doesn't exist and should be changed to`);
        });

        duplicates.forEach(gl => {
            const desc = masterGLMap.get(gl) ?? "(description not found)";
            body.push(`${gl} (${desc}) – error, duplicate`);
        });

        outputF[i][0] = `${ccmName}\n\n${parts.join(", ")}:\n${body.join("\n")}`;
    }

    ws.getRangeByIndexes(1, 5, rowCount, 1).setValues(outputF);
    ws.getRangeByIndexes(1, 33, rowCount, 1).setValues(outputAH);
}

function toProperCase(str: string): string {
    return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
}