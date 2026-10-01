import {
    assignSheetNames,
    assignStageChildKey,
    buildBreadcrumb,
    buildStageChildForm,
    buildTrackerRecords,
    computeColumnCount,
    crossCheckSurveyIds,
    emptyIntegrityReport,
    finalizeColumns,
    forecastWorkbook,
    formatBytes,
    keyColumnsFor,
    linkParents,
    pascalizeStageName,
    resolveFacilityIds,
} from "../extract-forms";
import type {
    FormData,
    FormDiscovery,
    ProgramMeta,
    ProgramStageMeta,
    Record_,
    ResolvedForm,
} from "../extract-forms";
import {
    PREVALENCE_CASE_REPORT_FORM_ID,
    PREVALENCE_FACILITY_LEVEL_FORM_ID,
    PREVALENCE_MORTALITY_FOLLOWUP_FORM,
    PREVALENCE_SURVEY_FORM_ID,
} from "../../data/entities/D2Survey";

function form(overrides: Partial<ResolvedForm> & Pick<ResolvedForm, "key">): ResolvedForm {
    return {
        requestedName: overrides.key,
        uid: `${overrides.key}-uid`,
        serverName: overrides.key,
        kind: "tracker",
        defaultUid: `${overrides.key}-uid`,
        isCustom: false,
        parentLinkField: "",
        surveyLinkField: "",
        ...overrides,
    };
}

function record(overrides: Partial<Record_> & Pick<Record_, "id">): Record_ {
    return {
        parentId: "",
        surveyId: "",
        facilityId: "",
        orgUnit: "OU1",
        label: overrides.id,
        values: new Map(),
        meta: new Map(),
        ...overrides,
    };
}

function stage(overrides: Partial<ProgramStageMeta> & Pick<ProgramStageMeta, "id" | "name">): ProgramStageMeta {
    return {
        repeatable: false,
        sortOrder: 0,
        dataElementLabels: new Map(),
        dataElementOrder: new Map(),
        dataElementSection: new Map(),
        ...overrides,
    };
}

function programMeta(overrides: Partial<ProgramMeta> & { stages?: ProgramStageMeta[] } = {}): ProgramMeta {
    const { stages, ...rest } = overrides;
    return {
        attributes: new Map(),
        attributeOrder: new Map(),
        dataElements: new Map(),
        stageById: new Map((stages ?? []).map(s => [s.id, s])),
        ...rest,
    };
}

function indexOf(all: FormData[]) {
    return {
        byKey: new Map(all.map(d => [d.form.key, d])),
        recordIndex: new Map(all.map(d => [d.form.key, new Map(d.records.map(r => [r.id, r]))])),
    };
}

/** The real graph: Facility and CaseReport both hang off Survey; leaves hang off CaseReport. */
function realGraphForms() {
    return {
        survey: form({ key: "Survey", uid: PREVALENCE_SURVEY_FORM_ID, defaultUid: PREVALENCE_SURVEY_FORM_ID, kind: "event" }),
        facility: form({
            key: "Facility",
            uid: PREVALENCE_FACILITY_LEVEL_FORM_ID,
            defaultUid: PREVALENCE_FACILITY_LEVEL_FORM_ID,
            parentKey: "Survey",
        }),
        caseReport: form({
            key: "CaseReport",
            uid: PREVALENCE_CASE_REPORT_FORM_ID,
            defaultUid: PREVALENCE_CASE_REPORT_FORM_ID,
            parentKey: "Survey",
        }),
        followUp: form({
            key: "FollowUpD28",
            uid: PREVALENCE_MORTALITY_FOLLOWUP_FORM,
            defaultUid: PREVALENCE_MORTALITY_FOLLOWUP_FORM,
            parentKey: "CaseReport",
        }),
    };
}

describe("linkParents", () => {
    it("links Case report to Survey, not Facility (the app filters it by parentSurveyId)", () => {
        const resolved = [
            form({ key: "Survey", uid: PREVALENCE_SURVEY_FORM_ID, defaultUid: PREVALENCE_SURVEY_FORM_ID }),
            form({ key: "Facility", uid: PREVALENCE_FACILITY_LEVEL_FORM_ID, defaultUid: PREVALENCE_FACILITY_LEVEL_FORM_ID }),
            // Deliberately wrong on the way in — linkParents must correct it.
            form({
                key: "CaseReport",
                uid: PREVALENCE_CASE_REPORT_FORM_ID,
                defaultUid: PREVALENCE_CASE_REPORT_FORM_ID,
                parentKey: "Facility",
            }),
        ];

        linkParents(resolved);

        expect(resolved.find(f => f.key === "CaseReport")?.parentKey).toBe("Survey");
        expect(resolved.find(f => f.key === "Facility")?.parentKey).toBe("Survey");
    });

    it("links a custom leaf form to Case report via the default UID it overrides", () => {
        const resolved = [
            form({ key: "CaseReport", uid: "custom-crf", defaultUid: PREVALENCE_CASE_REPORT_FORM_ID, isCustom: true }),
            form({
                key: "FollowUpD28",
                uid: "custom-followup",
                defaultUid: PREVALENCE_MORTALITY_FOLLOWUP_FORM,
                isCustom: true,
            }),
        ];

        linkParents(resolved);

        expect(resolved.find(f => f.key === "FollowUpD28")?.parentKey).toBe("CaseReport");
    });
});

describe("keyColumnsFor", () => {
    it("gives a leaf form Survey_id, Facility_id and CaseReport_id", () => {
        const { survey, facility, caseReport, followUp } = realGraphForms();
        const all: FormData[] = [survey, facility, caseReport, followUp].map(f => ({
            form: f,
            records: [],
            columns: new Map(),
        }));
        const { byKey } = indexOf(all);

        expect(keyColumnsFor(followUp, byKey).map(c => c.header)).toEqual([
            "Survey_id",
            "Facility_id",
            "CaseReport_id",
        ]);
    });

    it("gives Case report Survey_id and Facility_id but no parent column (its parent IS Survey)", () => {
        const { survey, facility, caseReport } = realGraphForms();
        const all: FormData[] = [survey, facility, caseReport].map(f => ({
            form: f,
            records: [],
            columns: new Map(),
        }));
        const { byKey } = indexOf(all);

        expect(keyColumnsFor(caseReport, byKey).map(c => c.header)).toEqual(["Survey_id", "Facility_id"]);
    });

    it("gives Facility only Survey_id, and Survey none", () => {
        const { survey, facility } = realGraphForms();
        const all: FormData[] = [survey, facility].map(f => ({ form: f, records: [], columns: new Map() }));
        const { byKey } = indexOf(all);

        expect(keyColumnsFor(facility, byKey).map(c => c.header)).toEqual(["Survey_id"]);
        expect(keyColumnsFor(survey, byKey)).toEqual([]);
    });
});

describe("resolveFacilityIds", () => {
    it("derives Facility_id from (Survey_id, org_unit_id) since no FK exists", () => {
        const { survey, facility, caseReport } = realGraphForms();
        const all: FormData[] = [
            { form: survey, records: [record({ id: "S1", surveyId: "S1" })], columns: new Map() },
            {
                form: facility,
                records: [record({ id: "F1", surveyId: "S1", orgUnit: "OU_A" })],
                columns: new Map(),
            },
            {
                form: caseReport,
                records: [
                    record({ id: "C1", surveyId: "S1", orgUnit: "OU_A" }),
                    record({ id: "C2", surveyId: "S1", orgUnit: "OU_UNKNOWN" }),
                ],
                columns: new Map(),
            },
        ];
        const report = emptyIntegrityReport();

        resolveFacilityIds(all, "Facility", report);

        const cases = all[2]!.records;
        expect(cases[0]!.facilityId).toBe("F1");
        expect(cases[1]!.facilityId).toBe("");
        expect(report.unresolvedFacility).toEqual([{ formKey: "CaseReport", count: 1 }]);
    });

    it("does not match across surveys even when the org unit is the same", () => {
        const { survey, facility, caseReport } = realGraphForms();
        const all: FormData[] = [
            { form: survey, records: [], columns: new Map() },
            { form: facility, records: [record({ id: "F1", surveyId: "S1", orgUnit: "OU_A" })], columns: new Map() },
            { form: caseReport, records: [record({ id: "C1", surveyId: "S2", orgUnit: "OU_A" })], columns: new Map() },
        ];
        const report = emptyIntegrityReport();

        resolveFacilityIds(all, "Facility", report);

        expect(all[2]!.records[0]!.facilityId).toBe("");
    });

    it("reports (survey, orgUnit) pairs that match more than one Facility", () => {
        const { facility, caseReport } = realGraphForms();
        const all: FormData[] = [
            {
                form: facility,
                records: [
                    record({ id: "F1", surveyId: "S1", orgUnit: "OU_A" }),
                    record({ id: "F2", surveyId: "S1", orgUnit: "OU_A" }),
                ],
                columns: new Map(),
            },
            { form: caseReport, records: [record({ id: "C1", surveyId: "S1", orgUnit: "OU_A" })], columns: new Map() },
        ];
        const report = emptyIntegrityReport();

        resolveFacilityIds(all, "Facility", report);

        expect(report.ambiguousFacilities).toEqual([
            { surveyId: "S1", orgUnit: "OU_A", facilityIds: ["F1", "F2"] },
        ]);
        // Still deterministic rather than silently dropped.
        expect(all[1]!.records[0]!.facilityId).toBe("F1");
    });

    it("marks a Facility record as its own facility", () => {
        const { facility } = realGraphForms();
        const all: FormData[] = [
            { form: facility, records: [record({ id: "F1", surveyId: "S1", orgUnit: "OU_A" })], columns: new Map() },
        ];

        resolveFacilityIds(all, "Facility", emptyIntegrityReport());

        expect(all[0]!.records[0]!.facilityId).toBe("F1");
    });
});

describe("crossCheckSurveyIds", () => {
    it("flags a leaf whose own Survey_id disagrees with the Survey via its parent", () => {
        const { caseReport, followUp } = realGraphForms();
        const all: FormData[] = [
            { form: caseReport, records: [record({ id: "C1", surveyId: "S1" })], columns: new Map() },
            {
                form: followUp,
                records: [record({ id: "U1", surveyId: "S_WRONG", parentId: "C1" })],
                columns: new Map(),
            },
        ];
        const { recordIndex } = indexOf(all);
        const report = emptyIntegrityReport();

        crossCheckSurveyIds(all, recordIndex, report);

        expect(report.surveyMismatches).toEqual([
            { formKey: "FollowUpD28", recordId: "U1", own: "S_WRONG", viaParent: "S1" },
        ]);
    });

    it("is silent when the ids agree", () => {
        const { caseReport, followUp } = realGraphForms();
        const all: FormData[] = [
            { form: caseReport, records: [record({ id: "C1", surveyId: "S1" })], columns: new Map() },
            { form: followUp, records: [record({ id: "U1", surveyId: "S1", parentId: "C1" })], columns: new Map() },
        ];
        const { recordIndex } = indexOf(all);
        const report = emptyIntegrityReport();

        crossCheckSurveyIds(all, recordIndex, report);

        expect(report.surveyMismatches).toEqual([]);
    });
});

describe("buildBreadcrumb", () => {
    const { survey, caseReport, followUp } = realGraphForms();

    const surveyRec = record({ id: "S1", label: "Kenya PPS 2024", surveyId: "S1" });
    const caseRec = record({ id: "C1", label: "Patient 0187", parentId: "S1", surveyId: "S1" });
    const followUpRec = record({ id: "U1", label: "D28", parentId: "C1", surveyId: "S1" });

    const all: FormData[] = [
        { form: survey, records: [surveyRec], columns: new Map() },
        { form: caseReport, records: [caseRec], columns: new Map() },
        { form: followUp, records: [followUpRec], columns: new Map() },
    ];

    it("walks the FK chain from the root down to the record", () => {
        const { byKey, recordIndex } = indexOf(all);

        expect(buildBreadcrumb(followUp, followUpRec, byKey, recordIndex)).toBe(
            "Kenya PPS 2024 > Patient 0187 > D28"
        );
    });

    it("returns just the record's own label for a root record", () => {
        const { byKey, recordIndex } = indexOf(all);
        expect(buildBreadcrumb(survey, surveyRec, byKey, recordIndex)).toBe("Kenya PPS 2024");
    });

    it("falls back to the raw id and reports an orphan when the parent is missing", () => {
        const orphan = record({ id: "U2", label: "D28-orphan", parentId: "MISSING" });
        const { byKey, recordIndex } = indexOf(all);
        const report = emptyIntegrityReport();

        const path = buildBreadcrumb(followUp, orphan, byKey, recordIndex, report);

        expect(path).toBe("MISSING > D28-orphan");
        expect(report.orphans).toEqual([
            { formKey: "FollowUpD28", recordId: "U2", missingParentId: "MISSING" },
        ]);
    });

    it("stops cleanly when a record has no parent id", () => {
        const detached = record({ id: "U3", label: "D28-detached", parentId: "" });
        const { byKey, recordIndex } = indexOf(all);

        expect(buildBreadcrumb(followUp, detached, byKey, recordIndex)).toBe("D28-detached");
    });

    it("does not loop forever if the data contains a parent cycle", () => {
        const a = record({ id: "A", label: "A", parentId: "B" });
        const b = record({ id: "B", label: "B", parentId: "A" });
        const selfRef = form({ key: "Loop", parentKey: "Loop" });
        const loopData: FormData[] = [{ form: selfRef, records: [a, b], columns: new Map() }];
        const { byKey, recordIndex } = indexOf(loopData);

        expect(buildBreadcrumb(selfRef, a, byKey, recordIndex)).toContain("A");
    });
});

describe("buildTrackerRecords", () => {
    const caseReport = form({ key: "CaseReport", surveyLinkField: "surveyTea" });

    it("merges a non-repeatable stage's dataValues onto the main row, prefixed by stage name", () => {
        const meta = programMeta({
            stages: [
                stage({
                    id: "STAGE1",
                    name: "Diagnosis",
                    repeatable: false,
                    dataElementLabels: new Map([["de1", "Primary site"]]),
                }),
            ],
        });
        const columns = new Map<string, string>();
        const stageColumns = new Map<string, Map<string, string>>();

        const { main, stageRecords } = buildTrackerRecords(
            {
                trackedEntity: "C1",
                orgUnit: "OU_A",
                attributes: [{ attribute: "surveyTea", value: "S1" }],
                events: [
                    {
                        event: "E1",
                        programStage: "STAGE1",
                        dataValues: [{ dataElement: "de1", value: "Lung" }],
                    },
                ],
            },
            caseReport,
            meta,
            columns,
            stageColumns
        );

        expect(main.values.get("stage:de1")).toBe("Lung");
        expect(columns.get("stage:de1")).toBe("Diagnosis: Primary site");
        expect(main.surveyId).toBe("S1");
        expect(stageRecords).toEqual([]);
    });

    it("routes a repeatable stage's events to child records instead of the main row", () => {
        const meta = programMeta({
            stages: [
                stage({
                    id: "WARD",
                    name: "Ward data",
                    repeatable: true,
                    dataElementLabels: new Map([["de2", "Ward name"]]),
                }),
            ],
        });
        const columns = new Map<string, string>();
        const stageColumns = new Map<string, Map<string, string>>();

        const { main, stageRecords } = buildTrackerRecords(
            {
                trackedEntity: "F1",
                orgUnit: "OU_A",
                events: [
                    { event: "W1", programStage: "WARD", occurredAt: "2024-01-01", dataValues: [{ dataElement: "de2", value: "ICU" }] },
                    { event: "W2", programStage: "WARD", occurredAt: "2024-02-01", dataValues: [{ dataElement: "de2", value: "ER" }] },
                ],
            },
            form({ key: "Facility" }),
            meta,
            columns,
            stageColumns
        );

        // The repeatable stage must NOT pollute the main row's columns.
        expect(main.values.size).toBe(0);
        expect(columns.size).toBe(0);

        expect(stageRecords).toHaveLength(2);
        expect(stageRecords[0]).toMatchObject({ stageId: "WARD" });
        expect(stageRecords[0]!.record).toMatchObject({ id: "W1", parentId: "F1", label: "2024-01-01" });
        expect(stageRecords[1]!.record).toMatchObject({ id: "W2", parentId: "F1" });
        expect(stageRecords[0]!.record.values.get("de2")).toBe("ICU");
        expect(stageRecords[1]!.record.values.get("de2")).toBe("ER");
        expect(stageColumns.get("WARD")?.get("de2")).toBe("Ward name");
    });

    it("copies the main row's surveyId onto repeatable-stage child records", () => {
        const meta = programMeta({ stages: [stage({ id: "WARD", name: "Ward", repeatable: true })] });

        const { stageRecords } = buildTrackerRecords(
            {
                trackedEntity: "F1",
                attributes: [{ attribute: "surveyTea", value: "S9" }],
                events: [{ event: "W1", programStage: "WARD", dataValues: [] }],
            },
            form({ key: "Facility", surveyLinkField: "surveyTea" }),
            meta,
            new Map(),
            new Map()
        );

        expect(stageRecords[0]!.record.surveyId).toBe("S9");
    });

    it("ignores an event on an unknown program stage rather than throwing", () => {
        const meta = programMeta({ stages: [] });

        expect(() =>
            buildTrackerRecords(
                {
                    trackedEntity: "C1",
                    events: [{ event: "E1", programStage: "UNKNOWN", dataValues: [{ dataElement: "de1", value: "x" }] }],
                },
                caseReport,
                meta,
                new Map(),
                new Map()
            )
        ).not.toThrow();
    });

    it("skips a dataValue with an undefined or null value", () => {
        const meta = programMeta({
            stages: [stage({ id: "S1", name: "S1", repeatable: false, dataElementLabels: new Map([["de1", "X"]]) })],
        });
        const columns = new Map<string, string>();

        const { main } = buildTrackerRecords(
            {
                trackedEntity: "C1",
                events: [{ event: "E1", programStage: "S1", dataValues: [{ dataElement: "de1", value: null }] }],
            },
            caseReport,
            meta,
            columns,
            new Map()
        );

        expect(main.values.has("stage:de1")).toBe(false);
        expect(columns.size).toBe(0);
    });
});

describe("keyColumnsFor (repeatable-stage children)", () => {
    it("does not duplicate Facility_id when a repeatable stage belongs to Facility itself", () => {
        const facility = form({ key: "Facility" });
        const wardStage = form({ key: "Facility_WardData", kind: "trackerStage", parentKey: "Facility" });
        const all: FormData[] = [
            { form: facility, records: [], columns: new Map() },
            { form: wardStage, records: [], columns: new Map() },
        ];
        const { byKey } = indexOf(all);

        const headers = keyColumnsFor(wardStage, byKey).map(c => c.header);

        expect(headers).toEqual(["Survey_id", "Facility_id"]);
        expect(headers.filter(h => h === "Facility_id")).toHaveLength(1);
    });

    it("gives a repeatable stage under Case report Survey_id, Facility_id and CaseReport_id", () => {
        const { survey, facility, caseReport } = realGraphForms();
        const stageForm = form({ key: "CaseReport_Diagnosis", kind: "trackerStage", parentKey: "CaseReport" });
        const all: FormData[] = [survey, facility, caseReport, stageForm].map(f => ({
            form: f,
            records: [],
            columns: new Map(),
        }));
        const { byKey } = indexOf(all);

        expect(keyColumnsFor(stageForm, byKey).map(c => c.header)).toEqual([
            "Survey_id",
            "Facility_id",
            "CaseReport_id",
        ]);
    });
});

describe("pascalizeStageName", () => {
    it("preserves word boundaries instead of blindly stripping separators", () => {
        expect(pascalizeStageName("Ward data")).toBe("WardData");
    });

    it("handles multiple words and punctuation", () => {
        expect(pascalizeStageName("SC/ID/AST point without gram stain test")).toBe(
            "SCIDASTPointWithoutGramStainTest"
        );
    });

    it("falls back to a stable placeholder for an empty/unusable name", () => {
        expect(pascalizeStageName("")).toBe("Stage");
        expect(pascalizeStageName("---")).toBe("Stage");
    });

    it("is idempotent-ish: already-PascalCase input round-trips sensibly", () => {
        expect(pascalizeStageName("Species 1")).toBe("Species1");
    });
});

describe("assignSheetNames", () => {
    it("proves the real collision case: two stage names that share a 31-char prefix no longer collide", () => {
        // This is the exact pair verified live against the server: two SampleShipment
        // stage names differing only after character 31.
        const sources = [
            { key: "SampleShipment__SCIDASTPointWithoutGramStainTest", stableId: "uid-A" },
            { key: "SampleShipment__SCIDASTPointWithGramStainTest", stableId: "uid-B" },
        ];

        const names = assignSheetNames(sources);

        const a = names.get(sources[0]!.key)!;
        const b = names.get(sources[1]!.key)!;
        expect(a).not.toBe(b);
        expect(a.length).toBeLessThanOrEqual(31);
        expect(b.length).toBeLessThanOrEqual(31);
    });

    it("never exceeds 31 characters even for a very long single key", () => {
        const names = assignSheetNames([{ key: "A".repeat(60), stableId: "uid-X" }]);
        expect([...names.values()][0]!.length).toBeLessThanOrEqual(31);
    });

    it("is deterministic: the same inputs in a different array order produce the same assignment", () => {
        const sources = [
            { key: "Zeta", stableId: "z" },
            { key: "Alpha", stableId: "a" },
            { key: "Beta", stableId: "b" },
        ];

        const forward = assignSheetNames(sources);
        const shuffled = assignSheetNames([...sources].reverse());

        expect(forward.get("Alpha")).toBe(shuffled.get("Alpha"));
        expect(forward.get("Beta")).toBe(shuffled.get("Beta"));
        expect(forward.get("Zeta")).toBe(shuffled.get("Zeta"));
    });

    it("is stable across runs: disambiguation depends on stableId, not encounter order", () => {
        const base = "X".repeat(31); // two sources that collide once truncated
        const sources = [
            { key: `${base}One`, stableId: "programStageUidOne" },
            { key: `${base}Two`, stableId: "programStageUidTwo" },
        ];

        const run1 = assignSheetNames(sources);
        const run2 = assignSheetNames([...sources]); // fresh call, same logical inputs

        expect(run1.get(sources[0]!.key)).toBe(run2.get(sources[0]!.key));
        expect(run1.get(sources[1]!.key)).toBe(run2.get(sources[1]!.key));
    });

    it("never assigns a name that collides with a reserved name", () => {
        const names = assignSheetNames([{ key: "_index", stableId: "weird-form-uid" }], ["_index", "_relationships"]);
        expect(names.get("_index")).not.toBe("_index");
    });

    it("records collisions in the given IntegrityReport instead of passing silently", () => {
        const report = emptyIntegrityReport();
        const base = "Y".repeat(31);
        assignSheetNames(
            [
                { key: `${base}A`, stableId: "uidA" },
                { key: `${base}B`, stableId: "uidB" },
            ],
            [],
            report
        );
        expect(report.sheetNameCollisions).toHaveLength(1);
    });
});

describe("assignStageChildKey", () => {
    const repeatable = (id: string, name: string) => stage({ id, name, repeatable: true });

    it("builds a readable, word-boundary-preserving key", () => {
        const used = new Set<string>(["Facility"]);
        expect(assignStageChildKey("Facility", repeatable("WARD", "Ward data"), used)).toBe("Facility__WardData");
    });

    it("disambiguates two stages that pascalize to the same name using the stage id, deterministically", () => {
        const used = new Set<string>(["Facility"]);
        const first = assignStageChildKey("Facility", repeatable("STAGE_A", "Ward Data"), used);
        const second = assignStageChildKey("Facility", repeatable("STAGE_B", "ward-data"), used);

        expect(first).toBe("Facility__WardData");
        expect(second).toBe("Facility__WardData__STAGE_B");
        expect(first).not.toBe(second);
    });
});

describe("buildStageChildForm", () => {
    it("produces a form shaped correctly for a repeatable-stage child sheet", () => {
        const owner = form({ key: "Facility", uid: "FAC_UID", defaultUid: "FAC_DEFAULT" });
        const wardStage = stage({ id: "WARD", name: "Ward data", repeatable: true });

        const child = buildStageChildForm(owner, wardStage, "Facility__WardData");

        expect(child.key).toBe("Facility__WardData");
        expect(child.parentKey).toBe("Facility");
        expect(child.kind).toBe("trackerStage");
        expect(child.uid).toBe("FAC_UID#WARD");
        expect(child.defaultUid).toBe("FAC_DEFAULT#WARD");
        expect(child.parentLinkField).toBe("");
    });
});

describe("column order and header uniqueness", () => {
    /**
     * Mirrors the real Central Ref Lab shape that caused the reported bug: a stage that
     * repeats a 4-field block per antibiotic, every block reusing the same field names,
     * with the form's own section names ("S1 - Antibiotic 1") as the only thing telling
     * them apart.
     */
    function speciesStage(): ProgramStageMeta {
        const labels = new Map<string, string>();
        const order = new Map<string, number>();
        const section = new Map<string, string>();

        labels.set("species", "Specify the species");
        order.set("species", 0);
        section.set("species", "Species 1");

        // Two antibiotic blocks, each (antibiotic, AST result).
        for (const slot of [1, 2]) {
            const abx = `abx${slot}`;
            const ast = `ast${slot}`;
            labels.set(abx, "Specify the antibiotic");
            labels.set(ast, "AST results");
            order.set(abx, slot * 10);
            order.set(ast, slot * 10 + 1);
            section.set(abx, `S1 - Antibiotic ${slot}`);
            section.set(ast, `S1 - Antibiotic ${slot}`);
        }

        return stage({
            id: "SPECIES1",
            name: "Species 1",
            sortOrder: 0,
            dataElementLabels: labels,
            dataElementOrder: order,
            dataElementSection: section,
        });
    }

    it("puts an antibiotic next to its OWN AST result, in form order", () => {
        const meta = programMeta({ stages: [speciesStage()] });
        const columns = new Map<string, string>();

        // Deliberately encountered out of form order, as paging would produce.
        buildTrackerRecords(
            {
                trackedEntity: "C1",
                events: [
                    {
                        event: "E1",
                        programStage: "SPECIES1",
                        dataValues: [
                            { dataElement: "ast2", value: "S" },
                            { dataElement: "abx1", value: "Teicoplanin" },
                            { dataElement: "species", value: "S. aureus" },
                            { dataElement: "ast1", value: "R" },
                            { dataElement: "abx2", value: "Ceftriaxone" },
                        ],
                    },
                ],
            },
            form({ key: "CentralRefLab" }),
            meta,
            columns,
            new Map()
        );

        const ordered = finalizeColumns(columns, meta);

        expect([...ordered.values()]).toEqual([
            "Species 1: Specify the species",
            "S1 - Antibiotic 1: Specify the antibiotic",
            "S1 - Antibiotic 1: AST results",
            "S1 - Antibiotic 2: Specify the antibiotic",
            "S1 - Antibiotic 2: AST results",
        ]);
    });

    it("keeps each column bound to its own value after reordering", () => {
        const meta = programMeta({ stages: [speciesStage()] });
        const columns = new Map<string, string>();

        const { main } = buildTrackerRecords(
            {
                trackedEntity: "C1",
                events: [
                    {
                        event: "E1",
                        programStage: "SPECIES1",
                        dataValues: [
                            { dataElement: "abx1", value: "Teicoplanin" },
                            { dataElement: "ast1", value: "R" },
                            { dataElement: "abx2", value: "Ceftriaxone" },
                            { dataElement: "ast2", value: "S" },
                        ],
                    },
                ],
            },
            form({ key: "CentralRefLab" }),
            meta,
            columns,
            new Map()
        );

        const ordered = finalizeColumns(columns, meta);
        const row = [...ordered.keys()].map(key => main.values.get(key) ?? "");

        // Antibiotic 1 is R; antibiotic 2 is S. Neither drug has two AST results.
        expect(row).toEqual(["Teicoplanin", "R", "Ceftriaxone", "S"]);
    });

    it("numbers repeats in form order when sections do not disambiguate them", () => {
        const labels = new Map([
            ["a", "Please, specify"],
            ["b", "Please, specify"],
        ]);
        const order = new Map([
            ["a", 1],
            ["b", 2],
        ]);
        const meta = programMeta({
            stages: [stage({ id: "S", name: "Ward data", dataElementLabels: labels, dataElementOrder: order })],
        });

        const columns = new Map([
            ["b", "Ward data: Please, specify"],
            ["a", "Ward data: Please, specify"],
        ]);

        expect([...finalizeColumns(columns, meta).entries()]).toEqual([
            ["a", "Ward data: Please, specify #1"],
            ["b", "Ward data: Please, specify #2"],
        ]);
    });

    it("leaves a label alone when it is already unique", () => {
        const meta = programMeta({
            stages: [
                stage({
                    id: "S",
                    name: "Stage",
                    dataElementLabels: new Map([["a", "Ward name"]]),
                    dataElementOrder: new Map([["a", 0]]),
                }),
            ],
        });

        const ordered = finalizeColumns(new Map([["a", "Stage: Ward name"]]), meta);
        expect([...ordered.values()]).toEqual(["Stage: Ward name"]);
    });

    it("orders attributes before stage data, and keeps unknown columns last", () => {
        const meta = programMeta({
            attributes: new Map([["attr", "Patient id"]]),
            attributeOrder: new Map([["attr", 0]]),
            stages: [
                stage({
                    id: "S",
                    name: "Stage",
                    dataElementLabels: new Map([["de", "Field"]]),
                    dataElementOrder: new Map([["de", 0]]),
                }),
            ],
        });

        const columns = new Map([
            ["orphaned", "Stage: Retired field"],
            ["stage:de", "Stage: Field"],
            ["attr", "Patient id"],
        ]);

        expect([...finalizeColumns(columns, meta).keys()]).toEqual(["attr", "stage:de", "orphaned"]);
    });
});

describe("buildTrackerRecords: events across all enrolments", () => {
    it("keeps events from every enrolment, not just the first", () => {
        // fetchTracker flattens enrollments[].events before calling this, so a TEI with a
        // second enrolment must not lose that enrolment's events from the extract.
        const meta = programMeta({
            stages: [stage({ id: "WARD", name: "Ward data", repeatable: true, dataElementLabels: new Map() })],
        });
        const stageColumns = new Map<string, Map<string, string>>();

        const { stageRecords } = buildTrackerRecords(
            {
                trackedEntity: "F1",
                enrollment: { enrolledAt: "2024-01-01", status: "ACTIVE" },
                events: [
                    { event: "W1", programStage: "WARD", dataValues: [] },
                    { event: "W2-from-second-enrolment", programStage: "WARD", dataValues: [] },
                ],
            },
            form({ key: "Facility" }),
            meta,
            new Map(),
            stageColumns
        );

        expect(stageRecords.map(s => s.record.id)).toEqual(["W1", "W2-from-second-enrolment"]);
    });

    it("takes enrolment-level metadata from the first enrolment", () => {
        const { main } = buildTrackerRecords(
            {
                trackedEntity: "F1",
                enrollment: { enrolledAt: "2024-01-01", status: "COMPLETED" },
                events: [],
            },
            form({ key: "Facility" }),
            programMeta(),
            new Map(),
            new Map()
        );

        expect(main.meta.get("enrolled_at")).toBe("2024-01-01");
        expect(main.meta.get("status")).toBe("COMPLETED");
    });
});

describe("buildTrackerRecords: unresolved-stage diagnostics", () => {
    it("counts (not silently drops) an event on an unknown program stage", () => {
        const meta = programMeta({ stages: [] });
        const report = emptyIntegrityReport();

        buildTrackerRecords(
            {
                trackedEntity: "C1",
                events: [{ event: "E1", programStage: "UNKNOWN", dataValues: [{ dataElement: "de1", value: "x" }] }],
            },
            form({ key: "CaseReport" }),
            meta,
            new Map(),
            new Map(),
            report
        );

        expect(report.unresolvedStageEvents).toEqual([{ formKey: "CaseReport", programStage: "UNKNOWN", count: 1 }]);
    });

    it("does not double-count the same unresolved event across the two internal passes", () => {
        const meta = programMeta({ stages: [] });
        const report = emptyIntegrityReport();

        buildTrackerRecords(
            {
                trackedEntity: "C1",
                events: [{ event: "E1", programStage: "UNKNOWN", dataValues: [] }],
            },
            form({ key: "CaseReport" }),
            meta,
            new Map(),
            new Map(),
            report
        );

        expect(report.unresolvedStageEvents[0]!.count).toBe(1);
    });

    it("accumulates counts across multiple calls for the same form and stage", () => {
        const meta = programMeta({ stages: [] });
        const report = emptyIntegrityReport();
        const teiInput = (id: string) => ({
            trackedEntity: id,
            events: [{ event: `E-${id}`, programStage: "UNKNOWN", dataValues: [] }],
        });

        buildTrackerRecords(teiInput("C1"), form({ key: "CaseReport" }), meta, new Map(), new Map(), report);
        buildTrackerRecords(teiInput("C2"), form({ key: "CaseReport" }), meta, new Map(), new Map(), report);

        expect(report.unresolvedStageEvents).toEqual([{ formKey: "CaseReport", programStage: "UNKNOWN", count: 2 }]);
    });

    it("does not record anything when every stage resolves", () => {
        const meta = programMeta({ stages: [stage({ id: "S1", name: "S1" })] });
        const report = emptyIntegrityReport();

        buildTrackerRecords(
            { trackedEntity: "C1", events: [{ event: "E1", programStage: "S1", dataValues: [] }] },
            form({ key: "CaseReport" }),
            meta,
            new Map(),
            new Map(),
            report
        );

        expect(report.unresolvedStageEvents).toEqual([]);
    });
});

describe("buildTrackerRecords: header label resolution (tiers 1-2)", () => {
    it("resolves a stage dataElement's label from its own stage config (tier 1)", () => {
        const meta = programMeta({
            stages: [stage({ id: "S1", name: "Diagnosis", dataElementLabels: new Map([["de1", "Primary site"]]) })],
        });
        const columns = new Map<string, string>();

        buildTrackerRecords(
            { trackedEntity: "C1", events: [{ event: "E1", programStage: "S1", dataValues: [{ dataElement: "de1", value: "Lung" }] }] },
            form({ key: "CaseReport" }),
            meta,
            columns,
            new Map()
        );

        expect(columns.get("stage:de1")).toBe("Diagnosis: Primary site");
    });

    it("falls back to the program's flat dataElement list when the stage config lacks the label (tier 2)", () => {
        // Live-verified real scenario: a dataElement removed from a stage's current config
        // but still holding historical values, still present in the program's flat list.
        const meta = programMeta({
            stages: [stage({ id: "S1", name: "Diagnosis", dataElementLabels: new Map() })], // empty: not in stage config
            dataElements: new Map([["de1", "Rescued From Flat List"]]),
        });
        const columns = new Map<string, string>();
        const unresolvedHeaders = undefined; // must not be required for tier 2 to work

        buildTrackerRecords(
            { trackedEntity: "C1", events: [{ event: "E1", programStage: "S1", dataValues: [{ dataElement: "de1", value: "Lung" }] }] },
            form({ key: "CaseReport" }),
            meta,
            columns,
            new Map(),
            undefined,
            unresolvedHeaders
        );

        expect(columns.get("stage:de1")).toBe("Diagnosis: Rescued From Flat List");
    });

    it("resolves a repeatable-stage child column the same way (stage config, then flat list)", () => {
        const meta = programMeta({
            stages: [stage({ id: "WARD", name: "Ward data", repeatable: true, dataElementLabels: new Map() })],
            dataElements: new Map([["de2", "Rescued Ward Field"]]),
        });
        const stageColumns = new Map<string, Map<string, string>>();

        buildTrackerRecords(
            {
                trackedEntity: "F1",
                events: [{ event: "W1", programStage: "WARD", dataValues: [{ dataElement: "de2", value: "ICU" }] }],
            },
            form({ key: "Facility" }),
            meta,
            new Map(),
            stageColumns
        );

        expect(stageColumns.get("WARD")?.get("de2")).toBe("Rescued Ward Field");
    });

    it("resolves an attribute label from meta.attributes", () => {
        const meta = programMeta({ attributes: new Map([["attr1", "Hospital name"]]) });
        const columns = new Map<string, string>();

        buildTrackerRecords(
            { trackedEntity: "C1", attributes: [{ attribute: "attr1", value: "Nairobi" }] },
            form({ key: "Facility" }),
            meta,
            columns,
            new Map()
        );

        expect(columns.get("attr1")).toBe("Hospital name");
    });
});

describe("buildTrackerRecords: header resolution tier-3 queueing", () => {
    it("queues a mainStageDataElement patch when both tier 1 and 2 fail, and uses the raw id as a placeholder label", () => {
        const meta = programMeta({ stages: [stage({ id: "S1", name: "Diagnosis", dataElementLabels: new Map() })] });
        const columns = new Map<string, string>();
        const tracker = { dataElementIds: new Set<string>(), attributeIds: new Set<string>(), patches: [] as any[] };

        buildTrackerRecords(
            { trackedEntity: "C1", events: [{ event: "E1", programStage: "S1", dataValues: [{ dataElement: "strayDE", value: "x" }] }] },
            form({ key: "CaseReport" }),
            meta,
            columns,
            new Map(),
            undefined,
            tracker
        );

        expect(columns.get("stage:strayDE")).toBe("Diagnosis: strayDE"); // placeholder until tier 3 runs
        expect(tracker.dataElementIds.has("strayDE")).toBe(true);
        expect(tracker.patches).toEqual([{ kind: "mainStageDataElement", id: "strayDE", stageName: "Diagnosis" }]);
    });

    it("queues a childStageDataElement patch (with the stage id, not the stage name) for a repeatable stage", () => {
        const meta = programMeta({ stages: [stage({ id: "WARD", name: "Ward data", repeatable: true, dataElementLabels: new Map() })] });
        const stageColumns = new Map<string, Map<string, string>>();
        const tracker = { dataElementIds: new Set<string>(), attributeIds: new Set<string>(), patches: [] as any[] };

        buildTrackerRecords(
            { trackedEntity: "F1", events: [{ event: "W1", programStage: "WARD", dataValues: [{ dataElement: "strayDE2", value: "x" }] }] },
            form({ key: "Facility" }),
            meta,
            new Map(),
            stageColumns,
            undefined,
            tracker
        );

        expect(stageColumns.get("WARD")?.get("strayDE2")).toBe("strayDE2");
        expect(tracker.patches).toEqual([{ kind: "childStageDataElement", id: "strayDE2", stageId: "WARD" }]);
    });

    it("queues an attribute patch when the attribute isn't in meta.attributes", () => {
        const meta = programMeta();
        const columns = new Map<string, string>();
        const tracker = { dataElementIds: new Set<string>(), attributeIds: new Set<string>(), patches: [] as any[] };

        buildTrackerRecords(
            { trackedEntity: "C1", attributes: [{ attribute: "strayAttr", value: "x" }] },
            form({ key: "Facility" }),
            meta,
            columns,
            new Map(),
            undefined,
            tracker
        );

        expect(columns.get("strayAttr")).toBe("strayAttr");
        expect(tracker.attributeIds.has("strayAttr")).toBe(true);
        expect(tracker.patches).toEqual([{ kind: "attribute", id: "strayAttr" }]);
    });

    it("queues nothing when no tracker is provided (optional, backward compatible)", () => {
        const meta = programMeta({ stages: [stage({ id: "S1", name: "Diagnosis", dataElementLabels: new Map() })] });

        expect(() =>
            buildTrackerRecords(
                { trackedEntity: "C1", events: [{ event: "E1", programStage: "S1", dataValues: [{ dataElement: "strayDE", value: "x" }] }] },
                form({ key: "CaseReport" }),
                meta,
                new Map(),
                new Map()
            )
        ).not.toThrow();
    });
});

describe("computeColumnCount", () => {
    it("matches the header shape writeFormSheet actually produces for a leaf form", () => {
        const { survey, facility, caseReport, followUp } = realGraphForms();
        const all: FormData[] = [survey, facility, caseReport, followUp].map(f => ({
            form: f,
            records: [],
            columns: new Map(),
        }));
        const byKey = new Map(all.map(d => [d.form.key, d]));

        // FollowUpD28: 3 key cols (Survey_id, Facility_id, CaseReport_id) + 4 fixed
        // (path, record_id, org_unit_id, org_unit_name) + metaKeyCount + valueColumnCount.
        const columns = computeColumnCount(followUp, byKey, 4, 6);
        expect(columns).toBe(3 + 4 + 4 + 6);
    });

    it("gives Survey (root, no ancestors) just fixed + meta + value columns", () => {
        const { survey } = realGraphForms();
        const byKey = new Map([[survey.key, { form: survey, records: [], columns: new Map() }]]);

        expect(computeColumnCount(survey, byKey, 4, 269)).toBe(0 + 4 + 4 + 269);
    });
});

describe("forecastWorkbook", () => {
    function discoveryFor(overrides: Partial<FormDiscovery> & { form: ResolvedForm }): FormDiscovery {
        return { stages: [], ...overrides };
    }

    it("computes total cells as the sum of rows x columns across main and stage sheets", () => {
        const discoveries: FormDiscovery[] = [
            discoveryFor({
                form: form({ key: "Facility" }),
                total: 40,
                columns: 51,
                stages: [
                    {
                        sheetKey: "Facility__WardData",
                        stageName: "Ward data",
                        parentFormKey: "Facility",
                        rows: 100,
                        columns: 230,
                    },
                ],
            }),
        ];

        const forecast = forecastWorkbook(discoveries);

        expect(forecast.totalCells).toBe(40 * 51 + 100 * 230);
        expect(forecast.sheets).toHaveLength(2);
    });

    it("warns on an unusually wide sheet", () => {
        const discoveries: FormDiscovery[] = [
            discoveryFor({ form: form({ key: "CentralRefLab" }), total: 806, columns: 518 }),
        ];

        const forecast = forecastWorkbook(discoveries);

        expect(forecast.warnings.some(w => w.includes("CentralRefLab"))).toBe(true);
    });

    it("does not warn for a modestly sized sheet", () => {
        const discoveries: FormDiscovery[] = [discoveryFor({ form: form({ key: "Facility" }), total: 40, columns: 51 })];

        expect(forecastWorkbook(discoveries).warnings).toEqual([]);
    });

    it("excludes a form proven to have 0 records — extraction skips it, so no sheet is produced", () => {
        // Live-verified case: DischargeEconomic had total=0, so extract() never calls
        // fetchTracker for it, so its (also-0-row) repeatable stage never materializes
        // either. The forecast must match, not merely list everything discovery saw.
        const discoveries: FormDiscovery[] = [
            discoveryFor({
                form: form({ key: "DischargeEconomic" }),
                total: 0,
                columns: 16,
                stages: [
                    {
                        sheetKey: "DischargeEconomic__EconomicEvaluation",
                        stageName: "Economic evaluation",
                        parentFormKey: "DischargeEconomic",
                        rows: 0,
                        columns: 64,
                    },
                ],
            }),
            discoveryFor({ form: form({ key: "Facility" }), total: 40, columns: 51 }),
        ];

        const forecast = forecastWorkbook(discoveries);

        expect(forecast.sheets.map(s => s.sheetKey)).toEqual(["Facility"]);
    });

    it("excludes a repeatable stage with 0 rows even when its owning form has data", () => {
        const discoveries: FormDiscovery[] = [
            discoveryFor({
                form: form({ key: "Facility" }),
                total: 40,
                columns: 51,
                stages: [
                    { sheetKey: "Facility__EmptyStage", stageName: "Empty", parentFormKey: "Facility", rows: 0, columns: 10 },
                    { sheetKey: "Facility__WardData", stageName: "Ward data", parentFormKey: "Facility", rows: 5, columns: 230 },
                ],
            }),
        ];

        const forecast = forecastWorkbook(discoveries);

        expect(forecast.sheets.map(s => s.sheetKey)).toEqual(["Facility", "Facility__WardData"]);
    });

    it("leaves cells undefined (not zero) when rows or columns aren't known, rather than under-forecasting", () => {
        const discoveries: FormDiscovery[] = [discoveryFor({ form: form({ key: "WardSummaryStats", kind: "dataSet" }) })];

        const forecast = forecastWorkbook(discoveries);

        expect(forecast.sheets[0]!.cells).toBeUndefined();
        expect(forecast.totalCells).toBe(0);
    });
});

describe("formatBytes", () => {
    it("renders bytes, MB and GB at sensible boundaries", () => {
        expect(formatBytes(500)).toBe("500 B");
        expect(formatBytes(5 * 1024 * 1024)).toBe("5.0 MB");
        expect(formatBytes(2 * 1024 * 1024 * 1024)).toBe("2.00 GB");
    });
});
