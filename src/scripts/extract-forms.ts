/**
 * Read-only extraction of AMR Surveys form data into a single Excel workbook.
 *
 * This script issues GET requests only. It never posts, updates or deletes.
 *
 * Form UIDs are resolved from live server metadata by name, because "custom"
 * forms are configured per parent survey in the DHIS2 datastore
 * (amr-surveys/modules) and therefore do not exist as constants in this repo.
 * The constants in data/entities/D2Survey.ts are used only as a cross-check.
 *
 * Usage:
 *   yarn extract-forms --dry-run
 *   yarn extract-forms --org-unit <uid> --start-date 2024-01-01 --end-date 2024-12-31
 */
import { command, run, string, boolean, flag, option, optional, number } from "cmd-ts";
import Excel from "exceljs";
import fs from "fs";
import path from "path";

import { D2Api } from "../types/d2-api";
import {
    createSessionManager,
    describeAuth,
    deriveEnvLabel,
    getD2APiFromInstance,
    getEnvVars,
    getInstance,
    warmUpSession,
    type SessionManager,
} from "./common";
import { AMRSurveyModule } from "../domain/entities/AMRSurveyModule";
import { getDefaultProgram } from "../data/utils/getDefaultProgram";
import { getParentDataElementForProgram } from "../data/utils/surveyProgramHelper";
import {
    AMR_SURVEYS_MORTALITY_TEA_SURVEY_ID_COH,
    AMR_SURVEYS_MORTALITY_TEA_SURVEY_ID_DEC,
    AMR_SURVEYS_MORTALITY_TEA_SURVEY_ID_DF,
    AMR_SURVEYS_MORTALITY_TEA_SURVEY_ID_FUP,
    AMR_SURVEYS_PREVALENCE_TEA_SURVEY_ID_CRF,
    AMR_SURVEYS_PREVALENCE_TEA_SURVEY_ID_CRL,
    AMR_SURVEYS_PREVALENCE_TEA_SURVEY_ID_PIS,
    AMR_SURVEYS_PREVALENCE_TEA_SURVEY_ID_SRL,
    AMR_SURVEYS_PREVALENCE_TEA_SURVEY_ID_SSTF,
    AMR_SURVEYS_PREVALENCE_TEA_UNIQUE_PATIENT_ID,
    SURVEY_ID_FACILITY_LEVEL_DATAELEMENT_ID,
    PREVALENCE_CASE_REPORT_FORM_ID,
    PREVALENCE_CENTRAL_REF_LAB_FORM_ID,
    PREVALENCE_FACILITY_LEVEL_FORM_ID,
    PREVALENCE_MORTALITY_COHORT_ENORL_FORM,
    PREVALENCE_MORTALITY_DISCHARGE_CLINICAL_FORM,
    PREVALENCE_MORTALITY_DISCHARGE_ECONOMIC_FORM,
    PREVALENCE_MORTALITY_FOLLOWUP_FORM,
    PREVALENCE_PATHOGEN_ISO_STORE_TRACK_ID,
    PREVALENCE_SAMPLE_SHIP_TRACK_FORM_ID,
    PREVALENCE_SUPRANATIONAL_REF_LAB_ID,
    PREVALENCE_SURVEY_FORM_ID,
    PREVALENCE_SURVEY_NAME_DATAELEMENT_ID,
    WARD_SUMMARY_STATISTICS_FORM_ID,
} from "../data/entities/D2Survey";

// --- Registry of the forms to extract -----------------------------------------

// "trackerStage" is synthetic: a repeatable program stage on a tracker program, split
// into its own sheet because a TEI can have many events per repeatable stage (a
// different cardinality than the flat, one-row-per-TEI model everything else uses).
type FormKind = "tracker" | "event" | "dataSet" | "trackerStage";

type FormSpec = {
    /** Short stable key; used as the Excel sheet name and in the FK columns. */
    key: string;
    /** Name as it appears on the DHIS2 server. Resolution is by this name. */
    requestedName: string;
    /** Default UID from D2Survey.ts, used only to cross-check the resolution. */
    expectedDefaultUid?: string;
    /** Key of the parent form in this registry, if any. */
    parentKey?: string;
    /** Forced kind; otherwise derived from server metadata. */
    kind?: FormKind;
};

const FORMS: FormSpec[] = [
    { key: "Survey", requestedName: "Survey form", expectedDefaultUid: PREVALENCE_SURVEY_FORM_ID },
    {
        key: "Facility",
        requestedName: "Facility-level form",
        expectedDefaultUid: PREVALENCE_FACILITY_LEVEL_FORM_ID,
        parentKey: "Survey",
    },
    {
        // Parent is Survey, NOT Facility — see the note on PARENT_BY_DEFAULT_UID.
        key: "CaseReport",
        requestedName: "Case report form - custom v1",
        expectedDefaultUid: PREVALENCE_CASE_REPORT_FORM_ID,
        parentKey: "Survey",
    },
    {
        key: "SampleShipment",
        requestedName: "Sample shipment and tracking form - custom v1",
        expectedDefaultUid: PREVALENCE_SAMPLE_SHIP_TRACK_FORM_ID,
        parentKey: "CaseReport",
    },
    {
        key: "CentralRefLab",
        requestedName: "Central reference laboratory ID/AST results form",
        expectedDefaultUid: PREVALENCE_CENTRAL_REF_LAB_FORM_ID,
        parentKey: "CaseReport",
    },
    {
        key: "PathogenIsolates",
        requestedName: "Pathogen Isolates storage and tracking log",
        expectedDefaultUid: PREVALENCE_PATHOGEN_ISO_STORE_TRACK_ID,
        parentKey: "CaseReport",
    },
    {
        key: "Supranational",
        requestedName: "Supranational Reference Laboratory ID/AST results form",
        expectedDefaultUid: PREVALENCE_SUPRANATIONAL_REF_LAB_ID,
        parentKey: "CaseReport",
    },
    {
        key: "FollowUpD28",
        requestedName: "Follow-up form D28 - custom v1",
        expectedDefaultUid: PREVALENCE_MORTALITY_FOLLOWUP_FORM,
        parentKey: "CaseReport",
    },
    // No default UID in the codebase: resolved purely by name, and its parent
    // link field is derived via the datastore reverse-map at runtime.
    { key: "Discharge", requestedName: "Discharge form", parentKey: "CaseReport" },
    {
        key: "DischargeClinical",
        requestedName: "Discharge form - Clinical Evaluation",
        expectedDefaultUid: PREVALENCE_MORTALITY_DISCHARGE_CLINICAL_FORM,
        parentKey: "CaseReport",
    },
    {
        key: "DischargeEconomic",
        requestedName: "Discharge form - Economical Evaluation",
        expectedDefaultUid: PREVALENCE_MORTALITY_DISCHARGE_ECONOMIC_FORM,
        parentKey: "CaseReport",
    },
    {
        key: "CohortEnrolment",
        requestedName: "Cohort 3 enrolment form - custom v1",
        expectedDefaultUid: PREVALENCE_MORTALITY_COHORT_ENORL_FORM,
        parentKey: "CaseReport",
    },
    {
        key: "WardSummaryStats",
        requestedName: "Ward Summary Statistics",
        expectedDefaultUid: WARD_SUMMARY_STATISTICS_FORM_ID,
        kind: "dataSet",
    },
];

// --- Types --------------------------------------------------------------------

export type ResolvedForm = FormSpec & {
    uid: string;
    serverName: string;
    kind: FormKind;
    /** UID of the default form this one overrides (equals uid when not custom). */
    defaultUid: string;
    isCustom: boolean;
    /** Attribute/data element UID on this form holding the parent's id. */
    parentLinkField: string;
    /** Attribute UID on this form holding the *root Survey* id ("" if none). */
    surveyLinkField: string;
};

/** One extracted record, normalised across tracker/event/dataSet. */
export type Record_ = {
    id: string;
    parentId: string;
    /** Root Survey id, read directly from this record's own survey-link attribute. */
    surveyId: string;
    /** Facility record id, derived from (surveyId, orgUnit) — not a stored FK. */
    facilityId: string;
    orgUnit: string;
    label: string;
    values: Map<string, string>;
    meta: Map<string, string>;
};

export type FormData = {
    form: ResolvedForm;
    records: Record_[];
    /** Ordered value column UIDs -> display label. */
    columns: Map<string, string>;
};

const EXCEL_MAX_ROWS = 1_048_576;

/** Matches glass-dev's bulkDownloadAMUFiles.ts, which this server tolerates well. */
const FETCH_CONCURRENCY = 6;

/**
 * Lower than FETCH_CONCURRENCY: a form fetch pulls full record pages (attributes + every
 * event + every data value), which is far heavier per request than a discovery probe.
 */
const EXTRACT_CONCURRENCY = 3;

/**
 * Runs `fn` over `items` with at most `limit` in flight, preserving input order in the
 * result. Bounded rather than unbounded so a wide form list can't open dozens of parallel
 * requests against DHIS2 at once.
 */
async function mapWithConcurrency<T, R>(
    items: readonly T[],
    limit: number,
    fn: (item: T) => Promise<R>
): Promise<R[]> {
    const results: R[] = [];
    for (let i = 0; i < items.length; i += limit) {
        results.push(...(await Promise.all(items.slice(i, i + limit).map(fn))));
    }
    return results;
}

/**
 * Parent form by *default* UID. Keyed by default UID so custom forms resolve to the
 * same parent as the default they override.
 *
 * IMPORTANT — this mirrors what the app actually filters on, which is NOT the visual
 * nesting of the UI. GetPaginatedSurveysUseCase.ts:49 does:
 *
 *   const parentId = isPrevalencePatientChild(type) ? parentPatientId : parentSurveyId;
 *
 * and isPrevalencePatientChild (PPSProgramsHelper.ts:239) covers only the 8 leaf forms.
 * So Facility AND Case report are both filtered by the *root Survey id*: Case report's
 * tlRPoWumrSa holds a Survey id, not a Facility id. Case report has no foreign key to
 * Facility at all — the UI scopes it by orgUnit (useSurveys.ts:100). Facility linkage is
 * therefore derived from (Survey_id, org_unit_id); see resolveFacilityIds().
 */
const PARENT_BY_DEFAULT_UID: Record<string, string> = {
    [PREVALENCE_FACILITY_LEVEL_FORM_ID]: PREVALENCE_SURVEY_FORM_ID,
    [PREVALENCE_CASE_REPORT_FORM_ID]: PREVALENCE_SURVEY_FORM_ID,
    [PREVALENCE_SAMPLE_SHIP_TRACK_FORM_ID]: PREVALENCE_CASE_REPORT_FORM_ID,
    [PREVALENCE_CENTRAL_REF_LAB_FORM_ID]: PREVALENCE_CASE_REPORT_FORM_ID,
    [PREVALENCE_PATHOGEN_ISO_STORE_TRACK_ID]: PREVALENCE_CASE_REPORT_FORM_ID,
    [PREVALENCE_SUPRANATIONAL_REF_LAB_ID]: PREVALENCE_CASE_REPORT_FORM_ID,
    [PREVALENCE_MORTALITY_FOLLOWUP_FORM]: PREVALENCE_CASE_REPORT_FORM_ID,
    [PREVALENCE_MORTALITY_DISCHARGE_CLINICAL_FORM]: PREVALENCE_CASE_REPORT_FORM_ID,
    [PREVALENCE_MORTALITY_DISCHARGE_ECONOMIC_FORM]: PREVALENCE_CASE_REPORT_FORM_ID,
    [PREVALENCE_MORTALITY_COHORT_ENORL_FORM]: PREVALENCE_CASE_REPORT_FORM_ID,
};

/**
 * The TEA on each form that holds the *root Survey* id, by default UID. Every leaf form
 * carries one of these in addition to its patient link (see parentPrevalenceSurveyIdList,
 * D2Survey.ts:143), so every record can be traced to its Survey directly, without walking
 * the chain — and the walked chain can be cross-checked against it.
 */
const SURVEY_LINK_BY_DEFAULT_UID: Record<string, string> = {
    [PREVALENCE_FACILITY_LEVEL_FORM_ID]: SURVEY_ID_FACILITY_LEVEL_DATAELEMENT_ID,
    [PREVALENCE_CASE_REPORT_FORM_ID]: AMR_SURVEYS_PREVALENCE_TEA_SURVEY_ID_CRF,
    [PREVALENCE_SAMPLE_SHIP_TRACK_FORM_ID]: AMR_SURVEYS_PREVALENCE_TEA_SURVEY_ID_SSTF,
    [PREVALENCE_CENTRAL_REF_LAB_FORM_ID]: AMR_SURVEYS_PREVALENCE_TEA_SURVEY_ID_CRL,
    [PREVALENCE_PATHOGEN_ISO_STORE_TRACK_ID]: AMR_SURVEYS_PREVALENCE_TEA_SURVEY_ID_PIS,
    [PREVALENCE_SUPRANATIONAL_REF_LAB_ID]: AMR_SURVEYS_PREVALENCE_TEA_SURVEY_ID_SRL,
    [PREVALENCE_MORTALITY_FOLLOWUP_FORM]: AMR_SURVEYS_MORTALITY_TEA_SURVEY_ID_FUP,
    [PREVALENCE_MORTALITY_DISCHARGE_CLINICAL_FORM]: AMR_SURVEYS_MORTALITY_TEA_SURVEY_ID_DF,
    [PREVALENCE_MORTALITY_DISCHARGE_ECONOMIC_FORM]: AMR_SURVEYS_MORTALITY_TEA_SURVEY_ID_DEC,
    [PREVALENCE_MORTALITY_COHORT_ENORL_FORM]: AMR_SURVEYS_MORTALITY_TEA_SURVEY_ID_COH,
};

// --- Connection ---------------------------------------------------------------

// Connection/auth is handled by ./common.ts, a port of the glass-dev script auth
// layer (PAT support, backend:"fetch", session warm-up, refresh + backoff).

// --- Resolution ---------------------------------------------------------------

async function fetchModules(api: D2Api): Promise<AMRSurveyModule[]> {
    try {
        const modules = await api
            .get<AMRSurveyModule[]>("/dataStore/amr-surveys/modules")
            .getData();
        return Array.isArray(modules) ? modules : [];
    } catch (err) {
        console.warn(
            "  ! Could not read datastore amr-surveys/modules; custom-form detection disabled."
        );
        return [];
    }
}

function normalise(name: string): string {
    return name.trim().toLowerCase().replace(/\s+/g, " ");
}

async function resolveForms(
    api: D2Api,
    modules: AMRSurveyModule[],
    wanted: FormSpec[]
): Promise<{ resolved: ResolvedForm[]; unresolved: { spec: FormSpec; reason: string }[] }> {
    const { programs, dataSets } = await api.metadata
        .get({
            programs: { fields: { id: true, name: true, programType: true } },
            dataSets: { fields: { id: true, name: true } },
        })
        .getData();

    const resolved: ResolvedForm[] = [];
    const unresolved: { spec: FormSpec; reason: string }[] = [];

    for (const spec of wanted) {
        const wantName = normalise(spec.requestedName);

        const programMatches = programs.filter(p => normalise(p.name) === wantName);
        const dataSetMatches = dataSets.filter(d => normalise(d.name) === wantName);
        const matches = [
            ...programMatches.map(p => ({
                id: p.id,
                name: p.name,
                kind: (p.programType === "WITH_REGISTRATION" ? "tracker" : "event") as FormKind,
            })),
            ...dataSetMatches.map(d => ({ id: d.id, name: d.name, kind: "dataSet" as FormKind })),
        ];

        if (matches.length === 0) {
            unresolved.push({ spec, reason: "no program or dataSet with this exact name" });
            continue;
        }
        if (matches.length > 1) {
            unresolved.push({
                spec,
                reason: `ambiguous, ${matches.length} matches: ${matches
                    .map(m => `${m.name} (${m.id})`)
                    .join("; ")}`,
            });
            continue;
        }

        const match = matches[0]!;
        const defaultUid = getDefaultProgram(match.id, modules);
        const parentLinkField =
            match.kind === "dataSet" ? "" : getParentDataElementForProgram(match.id, modules);

        resolved.push({
            ...spec,
            uid: match.id,
            serverName: match.name,
            kind: spec.kind ?? match.kind,
            defaultUid,
            isCustom: defaultUid !== match.id,
            parentLinkField,
            surveyLinkField: SURVEY_LINK_BY_DEFAULT_UID[defaultUid] ?? "",
        });
    }

    return { resolved, unresolved };
}

/** Re-derive parentKey from the resolved default UIDs, so custom forms slot in correctly. */
export function linkParents(resolved: ResolvedForm[]): void {
    const byDefaultUid = new Map(resolved.map(f => [f.defaultUid, f.key]));

    for (const form of resolved) {
        const parentDefaultUid = PARENT_BY_DEFAULT_UID[form.defaultUid];
        const derived = parentDefaultUid ? byDefaultUid.get(parentDefaultUid) : undefined;
        if (derived && derived !== form.parentKey) {
            console.warn(
                `  ! ${form.key}: parent corrected from '${form.parentKey}' to '${derived}' via datastore mapping`
            );
            form.parentKey = derived;
        }
    }
}

function reportResolution(resolved: ResolvedForm[], unresolved: { spec: FormSpec; reason: string }[]) {
    console.log("\nResolved forms:");
    console.log(
        "  " +
            ["KEY".padEnd(18), "UID".padEnd(13), "KIND".padEnd(8), "CUSTOM".padEnd(7), "NAME"].join("")
    );
    for (const f of resolved) {
        const drift =
            f.expectedDefaultUid && f.defaultUid !== f.expectedDefaultUid
                ? `  <-- default ${f.defaultUid} != expected ${f.expectedDefaultUid}`
                : "";
        console.log(
            "  " +
                [
                    f.key.padEnd(18),
                    f.uid.padEnd(13),
                    f.kind.padEnd(8),
                    (f.isCustom ? "yes" : "no").padEnd(7),
                    f.serverName,
                ].join("") +
                drift
        );
        if (f.kind !== "dataSet" && !f.parentLinkField && f.parentKey) {
            console.warn(
                `      ! no parent link field known for ${f.key}; rows will have an empty ${f.parentKey}_id`
            );
        }
    }

    if (unresolved.length > 0) {
        console.log("\nUnresolved forms:");
        for (const { spec, reason } of unresolved) {
            console.log(`  ${spec.key.padEnd(18)} "${spec.requestedName}" -> ${reason}`);
        }
    }
}

// --- Metadata for column labels ------------------------------------------------

export type ProgramStageMeta = {
    id: string;
    name: string;
    repeatable: boolean;
    /** Position of this stage within its program. */
    sortOrder: number;
    /** dataElement id -> display label, plain (not prefixed with the stage/section name). */
    dataElementLabels: Map<string, string>;
    /**
     * dataElement id -> its position in the form (programStageDataElements.sortOrder).
     * Without this, columns land in the order values happen to be encountered while paging,
     * which scatters fields that belong together — e.g. an antibiotic and its own AST
     * result ending up nine columns apart on the Central Ref Lab sheet.
     */
    dataElementOrder: Map<string, number>;
    /**
     * dataElement id -> the form section it sits in, where the stage defines sections.
     * These carry the form's own slot names ("S1 - Antibiotic 1"), which is what makes
     * otherwise-identical labels like "Specify the antibiotic" tellable apart.
     */
    dataElementSection: Map<string, string>;
};

export type ProgramMeta = {
    attributes: Map<string, string>;
    /** attribute id -> its position in the program, so attribute columns keep form order. */
    attributeOrder: Map<string, number>;
    /** Used by fetchEvents (event programs have no stages in the model here). */
    dataElements: Map<string, string>;
    stageById: Map<string, ProgramStageMeta>;
};

/**
 * Not every dataElement/attribute referenced by a real value is guaranteed to appear in
 * either its own stage's config or the program's flat dataElement list — verified live: a
 * dataElement can be unassigned from a stage yet still hold historical values, and absent
 * from the program's flat /programs/{id} dataElements too. A raw id in a header is a
 * completeness signal users cannot silently miss, so any id that resists resolution here
 * is queued for one direct-fetch rescue attempt rather than left as-is.
 */
export type PendingHeaderPatch =
    | { kind: "attribute"; id: string }
    | { kind: "mainStageDataElement"; id: string; stageName: string }
    | { kind: "childStageDataElement"; id: string; stageId: string };

export type UnresolvedHeaderTracker = {
    dataElementIds: Set<string>;
    attributeIds: Set<string>;
    patches: PendingHeaderPatch[];
};

export function newUnresolvedHeaderTracker(): UnresolvedHeaderTracker {
    return { dataElementIds: new Set(), attributeIds: new Set(), patches: [] };
}

/**
 * Program metadata is needed twice per form — once by discovery (to predict column counts
 * and find repeatable stages) and again by extraction (to label columns). These payloads
 * are large (one program here has ~500 data elements), so the result is memoised per run;
 * a run only ever fetches each program's metadata once.
 */
const programMetaCache = new Map<string, Promise<ProgramMeta>>();

function fetchProgramMeta(api: D2Api, programId: string): Promise<ProgramMeta> {
    const cached = programMetaCache.get(programId);
    if (cached) return cached;
    // Evict on failure, so a transient error doesn't poison every later caller with the
    // same rejected promise and defeat the surrounding retry.
    const pending = fetchProgramMetaUncached(api, programId).catch(err => {
        programMetaCache.delete(programId);
        throw err;
    });
    programMetaCache.set(programId, pending);
    return pending;
}

type D2ProgramStage = {
    id: string;
    name: string;
    repeatable?: boolean;
    sortOrder?: number;
    programStageDataElements?: {
        sortOrder?: number;
        dataElement: { id: string; name: string; formName?: string };
    }[];
    programStageSections?: { name: string; sortOrder?: number; dataElements?: { id: string }[] }[];
};

async function fetchProgramMetaUncached(api: D2Api, programId: string): Promise<ProgramMeta> {
    // Two calls, because /programs/{id}/metadata.json does not reliably honour a nested
    // field selector for programStages (verified: it returns stage dataElement ids but
    // drops their names, sortOrder and sections). /programStages.json does.
    const [program, stagesResp] = await Promise.all([
        api
            .get<{
                programs?: {
                    programTrackedEntityAttributes?: {
                        sortOrder?: number;
                        trackedEntityAttribute: { id: string; name: string; formName?: string };
                    }[];
                }[];
                dataElements?: { id: string; name: string; formName?: string }[];
                trackedEntityAttributes?: { id: string; name: string; formName?: string }[];
            }>(`/programs/${programId}/metadata.json`, {
                fields: "programs,dataElements,trackedEntityAttributes,programTrackedEntityAttributes",
            })
            .getData(),
        api
            .get<{ programStages?: D2ProgramStage[] }>("/programStages.json", {
                filter: `program.id:eq:${programId}`,
                fields:
                    "id,name,repeatable,sortOrder," +
                    "programStageDataElements[sortOrder,dataElement[id,name,formName]]," +
                    "programStageSections[name,sortOrder,dataElements[id]]",
                paging: false,
            })
            .getData(),
    ]);

    const attributes = new Map<string, string>();
    const attributeOrder = new Map<string, number>();
    for (const tea of program.trackedEntityAttributes ?? []) {
        attributes.set(tea.id, tea.formName || tea.name);
    }
    for (const p of program.programs ?? []) {
        (p.programTrackedEntityAttributes ?? []).forEach((ptea, index) => {
            const tea = ptea.trackedEntityAttribute;
            if (!attributes.has(tea.id)) attributes.set(tea.id, tea.formName || tea.name);
            attributeOrder.set(tea.id, ptea.sortOrder ?? index);
        });
    }

    const dataElements = new Map<string, string>();
    for (const de of program.dataElements ?? []) {
        dataElements.set(de.id, de.formName || de.name);
    }

    const stages = [...(stagesResp.programStages ?? [])].sort(
        (a, b) => (a.sortOrder ?? 0) - (b.sortOrder ?? 0)
    );

    const stageById = new Map<string, ProgramStageMeta>();
    stages.forEach((stage, stageIndex) => {
        const dataElementLabels = new Map<string, string>();
        const dataElementOrder = new Map<string, number>();
        (stage.programStageDataElements ?? []).forEach((psde, index) => {
            const de = psde.dataElement;
            dataElementLabels.set(de.id, de.formName || de.name);
            dataElementOrder.set(de.id, psde.sortOrder ?? index);
        });

        const dataElementSection = new Map<string, string>();
        for (const section of stage.programStageSections ?? []) {
            for (const de of section.dataElements ?? []) dataElementSection.set(de.id, section.name);
        }

        stageById.set(stage.id, {
            id: stage.id,
            name: stage.name,
            repeatable: !!stage.repeatable,
            sortOrder: stage.sortOrder ?? stageIndex,
            dataElementLabels,
            dataElementOrder,
            dataElementSection,
        });
    });

    return { attributes, attributeOrder, dataElements, stageById };
}

/**
 * The label a column gets, before de-duplication: the form's own section name where the
 * stage defines one (e.g. "S1 - Antibiotic 1"), otherwise the stage name.
 */
function stageColumnLabel(stage: ProgramStageMeta, dataElementId: string, fieldLabel: string): string {
    return `${stage.dataElementSection.get(dataElementId) ?? stage.name}: ${fieldLabel}`;
}

/**
 * Marks a main-sheet column as coming from a program-stage data element rather than a
 * tracked-entity attribute, so the two id spaces can share one column map.
 */
const STAGE_COLUMN_PREFIX = "stage:";

/** Where a column belongs in form order. Attributes first, then stage data in stage order. */
const STAGE_COLUMN_OFFSET = 1_000_000;

function columnSortKey(columnKey: string, meta: ProgramMeta): number {
    const isStageValue = columnKey.startsWith(STAGE_COLUMN_PREFIX);
    const id = isStageValue ? columnKey.slice(STAGE_COLUMN_PREFIX.length) : columnKey;

    if (!isStageValue) {
        const attrPosition = meta.attributeOrder.get(id);
        if (attrPosition !== undefined) return attrPosition;
    }
    for (const stage of meta.stageById.values()) {
        const position = stage.dataElementOrder.get(id);
        if (position !== undefined) {
            return STAGE_COLUMN_OFFSET * (1 + stage.sortOrder) + position;
        }
    }
    // Not in this program's metadata (e.g. a field since unassigned from its stage):
    // keep it after everything known, rather than guessing a position.
    return Number.MAX_SAFE_INTEGER;
}

/**
 * Puts columns into form order and guarantees every header is unique.
 *
 * Order matters for reading: the Central Ref Lab form repeats a 4-field block per
 * antibiotic, and without this the antibiotic and its own AST result do not end up
 * side by side. Uniqueness matters because a repeated block reuses the same field names
 * — a bare "Specify the antibiotic" appearing 24 times tells the reader nothing. Where a
 * label still repeats after section prefixing, occurrences are numbered in form order.
 */
export function finalizeColumns(columns: Map<string, string>, meta: ProgramMeta): Map<string, string> {
    const ordered = [...columns.entries()].sort(
        ([a], [b]) => columnSortKey(a, meta) - columnSortKey(b, meta)
    );

    const labelCounts = new Map<string, number>();
    for (const [, label] of ordered) labelCounts.set(label, (labelCounts.get(label) ?? 0) + 1);

    const seen = new Map<string, number>();
    const result = new Map<string, string>();
    for (const [key, label] of ordered) {
        if ((labelCounts.get(label) ?? 0) === 1) {
            result.set(key, label);
            continue;
        }
        const occurrence = (seen.get(label) ?? 0) + 1;
        seen.set(label, occurrence);
        result.set(key, `${label} #${occurrence}`);
    }
    return result;
}

// --- Extraction ---------------------------------------------------------------

type FetchOpts = {
    orgUnit?: string;
    pageSize: number;
    startDate?: string;
    endDate?: string;
};

/**
 * Meta-column-key lists, named once so the discovery/forecast column-count math (which
 * predicts a sheet's shape before any records are fetched) can never drift from what the
 * real fetch functions below actually write — both reference the same `.length`.
 */
const TRACKER_MAIN_META_KEYS = ["created_at", "updated_at", "enrolled_at", "status"] as const;
const TRACKER_STAGE_CHILD_META_KEYS = ["occurred_at", "status"] as const;
const EVENT_META_KEYS = ["occurred_at", "created_at", "updated_at", "status"] as const;
/** path, record_id, org_unit_id, org_unit_name — present on every sheet regardless of kind. */
const FIXED_NON_META_COLUMNS = 4;

/**
 * The ResolvedForm shape a repeatable stage's child sheet gets. Shared by fetchTracker
 * (which creates these for real) and discovery's forecast (which previews them before any
 * extraction runs), so a predicted sheet can never describe a shape extraction wouldn't
 * actually produce.
 */
export function buildStageChildForm(owner: ResolvedForm, stage: ProgramStageMeta, childKey: string): ResolvedForm {
    return {
        key: childKey,
        requestedName: stage.name,
        parentKey: owner.key,
        uid: `${owner.uid}#${stage.id}`,
        serverName: stage.name,
        kind: "trackerStage",
        defaultUid: `${owner.defaultUid}#${stage.id}`,
        isCustom: owner.isCustom,
        // Not a stored attribute: parentId is the owning TEI's own id, so there is no
        // attribute UID to record here.
        parentLinkField: "",
        surveyLinkField: "",
    };
}

/**
 * Deterministic, collision-safe internal key for a repeatable stage's child sheet. Given
 * the same owner + stage processed in the same (id-sorted) order, this always produces the
 * same key — which is what lets discovery's prediction match extraction's actual output.
 */
export function assignStageChildKey(ownerKey: string, stage: ProgramStageMeta, usedKeys: Set<string>): string {
    const baseKey = `${ownerKey}__${pascalizeStageName(stage.name)}`;
    const childKey = usedKeys.has(baseKey) ? `${baseKey}__${stage.id}` : baseKey;
    usedKeys.add(childKey);
    return childKey;
}

function ouParams(orgUnit: string | undefined) {
    // ACCESSIBLE returns everything the authenticated user may read, which is
    // what "all data" means here. A root org unit narrows it to that subtree.
    return orgUnit
        ? { orgUnit, ouMode: "DESCENDANTS" as const }
        : { ouMode: "ACCESSIBLE" as const };
}

/**
 * Guards the paging loops. A server that ignores the `page` parameter would
 * otherwise return the same page forever; bail out as soon as a page adds no
 * record we have not already seen.
 */
function makePageGuard(formKey: string) {
    const seen = new Set<string>();

    return function accept(ids: string[]): boolean {
        const added = ids.filter(id => !seen.has(id));
        added.forEach(id => seen.add(id));

        if (ids.length > 0 && added.length === 0) {
            console.warn(
                `\n  ! ${formKey}: a page returned no new records; stopping to avoid a paging loop.`
            );
            return false;
        }
        return true;
    };
}

// --- Program-stage event handling ----------------------------------------------
//
// A tracker TEI carries data in two places: tracked-entity attributes, and events on
// its enrollment's program stages. A non-repeatable stage has at most one event per
// TEI, so its dataValues can be flattened onto the same row as extra columns. A
// repeatable stage can have many events per TEI (e.g. multiple ward records per
// facility) — that is a different cardinality, so it becomes its own child sheet,
// one row per event, FK'd back to the owning TEI.

export type StageEventInput = {
    event: string;
    programStage: string;
    occurredAt?: string;
    status?: string;
    dataValues?: { dataElement: string; value?: string | number | boolean | null }[];
};

export type TrackedEntityInput = {
    trackedEntity: string;
    orgUnit?: string;
    createdAt?: string;
    updatedAt?: string;
    attributes?: { attribute: string; value?: string | number | boolean | null }[];
    /** Enrolment-level metadata, taken from the TEI's first enrolment. */
    enrollment?: { enrolledAt?: string; status?: string };
    /** Events from ALL of the TEI's enrolments — never just the first, or a second
     * enrolment's events would vanish from the extract without a trace. */
    events?: StageEventInput[];
};

export type StageChildRecord = { stageId: string; record: Record_ };

/**
 * Builds one TEI's main row (attributes + non-repeatable-stage dataValues merged in)
 * plus one child Record_ per repeatable-stage event. Mutates `columns` and
 * `stageColumns` (per stage id) with any newly-seen column, matching the accumulation
 * pattern the rest of this file uses.
 */
export function buildTrackerRecords(
    tei: TrackedEntityInput,
    form: ResolvedForm,
    meta: ProgramMeta,
    columns: Map<string, string>,
    stageColumns: Map<string, Map<string, string>>,
    report?: IntegrityReport,
    unresolvedHeaders?: UnresolvedHeaderTracker
): { main: Record_; stageRecords: StageChildRecord[] } {
    const values = new Map<string, string>();
    for (const attr of tei.attributes ?? []) {
        if (attr.value === undefined || attr.value === null) continue;
        values.set(attr.attribute, String(attr.value));
        if (!columns.has(attr.attribute)) {
            const label = meta.attributes.get(attr.attribute);
            columns.set(attr.attribute, label ?? attr.attribute);
            if (!label && unresolvedHeaders) {
                unresolvedHeaders.attributeIds.add(attr.attribute);
                unresolvedHeaders.patches.push({ kind: "attribute", id: attr.attribute });
            }
        }
    }

    const stageRecords: StageChildRecord[] = [];

    for (const event of tei.events ?? []) {
        const stage = meta.stageById.get(event.programStage);
        if (!stage) {
            // Unknown stage: this event's data cannot be classified as repeatable or not,
            // so it is skipped. Not silent — every occurrence is counted and surfaced in
            // _index (see recordUnresolvedStageEvent), never just dropped unnoticed.
            if (report) recordUnresolvedStageEvent(report, form.key, event.programStage);
            continue;
        }

        if (!stage.repeatable) {
            for (const dv of event.dataValues ?? []) {
                if (dv.value === undefined || dv.value === null) continue;
                const key = `${STAGE_COLUMN_PREFIX}${dv.dataElement}`;
                values.set(key, String(dv.value));
                if (!columns.has(key)) {
                    // Tier 1: this stage's own config. Tier 2: the program's flat dataElement
                    // list — catches a dataElement that has since been unassigned from this
                    // stage but still holds historical values (verified live).
                    const label = stage.dataElementLabels.get(dv.dataElement) ?? meta.dataElements.get(dv.dataElement);
                    columns.set(key, stageColumnLabel(stage, dv.dataElement, label ?? dv.dataElement));
                    if (!label && unresolvedHeaders) {
                        unresolvedHeaders.dataElementIds.add(dv.dataElement);
                        unresolvedHeaders.patches.push({
                            kind: "mainStageDataElement",
                            id: dv.dataElement,
                            stageName: stage.dataElementSection.get(dv.dataElement) ?? stage.name,
                        });
                    }
                }
            }
        }
    }

    const surveyId = form.surveyLinkField ? values.get(form.surveyLinkField) ?? "" : "";
    const main: Record_ = {
        id: tei.trackedEntity,
        parentId: form.parentLinkField ? values.get(form.parentLinkField) ?? "" : "",
        surveyId,
        facilityId: "", // derived later by resolveFacilityIds()
        orgUnit: tei.orgUnit ?? "",
        label: values.get(AMR_SURVEYS_PREVALENCE_TEA_UNIQUE_PATIENT_ID) ?? tei.trackedEntity,
        values,
        meta: new Map<string, string>([
            ["created_at", tei.createdAt ?? ""],
            ["updated_at", tei.updatedAt ?? ""],
            ["enrolled_at", tei.enrollment?.enrolledAt ?? ""],
            ["status", tei.enrollment?.status ?? ""],
        ]),
    };

    for (const event of tei.events ?? []) {
        const stage = meta.stageById.get(event.programStage);
        // An unresolved stage was already counted in the loop above (same event list);
        // don't record it twice here.
        if (!stage || !stage.repeatable) continue;

        const childValues = new Map<string, string>();
        let childColumns = stageColumns.get(stage.id);
        if (!childColumns) {
            childColumns = new Map<string, string>();
            stageColumns.set(stage.id, childColumns);
        }
        for (const dv of event.dataValues ?? []) {
            if (dv.value === undefined || dv.value === null) continue;
            childValues.set(dv.dataElement, String(dv.value));
            if (!childColumns.has(dv.dataElement)) {
                const label = stage.dataElementLabels.get(dv.dataElement) ?? meta.dataElements.get(dv.dataElement);
                childColumns.set(dv.dataElement, label ?? dv.dataElement);
                if (!label && unresolvedHeaders) {
                    unresolvedHeaders.dataElementIds.add(dv.dataElement);
                    unresolvedHeaders.patches.push({
                        kind: "childStageDataElement",
                        id: dv.dataElement,
                        stageId: stage.id,
                    });
                }
            }
        }

        stageRecords.push({
            stageId: stage.id,
            record: {
                id: event.event,
                parentId: tei.trackedEntity,
                surveyId,
                facilityId: "",
                orgUnit: tei.orgUnit ?? "",
                label: event.occurredAt ?? event.event,
                values: childValues,
                meta: new Map([
                    ["occurred_at", event.occurredAt ?? ""],
                    ["status", event.status ?? ""],
                ]),
            },
        });
    }

    return { main, stageRecords };
}

/** Batched direct fetch of dataElement display labels for an arbitrary id list. */
async function fetchDataElementLabels(
    api: D2Api,
    ids: string[],
    session: SessionManager
): Promise<Map<string, string>> {
    const labels = new Map<string, string>();
    if (ids.length === 0) return labels;
    const resp = await session.retryWithBackoff(() =>
        api
            .get<{ dataElements?: { id: string; name: string; formName?: string }[] }>("/dataElements.json", {
                filter: `id:in:[${ids.join(",")}]`,
                fields: "id,name,formName",
                paging: false,
            })
            .getData()
    );
    for (const de of resp.dataElements ?? []) labels.set(de.id, de.formName || de.name);
    return labels;
}

/** Batched direct fetch of tracked-entity-attribute display labels for an arbitrary id list. */
async function fetchAttributeLabels(
    api: D2Api,
    ids: string[],
    session: SessionManager
): Promise<Map<string, string>> {
    const labels = new Map<string, string>();
    if (ids.length === 0) return labels;
    const resp = await session.retryWithBackoff(() =>
        api
            .get<{ trackedEntityAttributes?: { id: string; name: string; formName?: string }[] }>(
                "/trackedEntityAttributes.json",
                { filter: `id:in:[${ids.join(",")}]`, fields: "id,name,formName", paging: false }
            )
            .getData()
    );
    for (const a of resp.trackedEntityAttributes ?? []) labels.set(a.id, a.formName || a.name);
    return labels;
}

/**
 * Tier 3: a single batched direct fetch for every id that resisted tiers 1-2 across an
 * entire form's extraction (not per-record — `columns.has(key)` guards mean each distinct
 * id only ever queues one patch), then rewrites the affected column labels in place. Any
 * id that STILL doesn't resolve — genuinely deleted from DHIS2, not merely unassigned from
 * a stage — is recorded in `report.unresolvedHeaders` rather than left silently as a UID.
 */
async function resolveStrayHeaderLabels(
    api: D2Api,
    formKey: string,
    tracker: UnresolvedHeaderTracker,
    columns: Map<string, string>,
    stageColumnsByStageId: Map<string, Map<string, string>>,
    session: SessionManager,
    report: IntegrityReport
): Promise<void> {
    if (tracker.patches.length === 0) return;

    const [dataElementLabels, attributeLabels] = await Promise.all([
        fetchDataElementLabels(api, [...tracker.dataElementIds], session),
        fetchAttributeLabels(api, [...tracker.attributeIds], session),
    ]);

    for (const patch of tracker.patches) {
        if (patch.kind === "attribute") {
            const label = attributeLabels.get(patch.id);
            if (label) columns.set(patch.id, label);
            else report.unresolvedHeaders.push({ formKey, id: patch.id, kind: "trackedEntityAttribute" });
        } else if (patch.kind === "mainStageDataElement") {
            const label = dataElementLabels.get(patch.id);
            if (label) columns.set(`${STAGE_COLUMN_PREFIX}${patch.id}`, `${patch.stageName}: ${label}`);
            else report.unresolvedHeaders.push({ formKey, id: patch.id, kind: "dataElement" });
        } else {
            const label = dataElementLabels.get(patch.id);
            const childColumns = stageColumnsByStageId.get(patch.stageId);
            if (label && childColumns) childColumns.set(patch.id, label);
            else if (!label) report.unresolvedHeaders.push({ formKey, id: patch.id, kind: "dataElement" });
        }
    }

    console.warn(
        `  ! ${formKey}: ${tracker.patches.length} column header(s) needed a direct-fetch rescue ` +
            `(stage config / program metadata didn't have the name)`
    );
}

/**
 * Word-boundary-preserving PascalCase fragment from a stage name, e.g.
 * "Ward data" -> "WardData" (not "Warddata" — blind non-alnum stripping loses the
 * boundary between words, which makes multi-word stage names unreadable).
 */
export function pascalizeStageName(name: string): string {
    const words = name.split(/[^a-zA-Z0-9]+/).filter(Boolean);
    const pascal = words.map(w => w.charAt(0).toUpperCase() + w.slice(1)).join("");
    return pascal || "Stage";
}

async function fetchTracker(
    api: D2Api,
    form: ResolvedForm,
    meta: ProgramMeta,
    opts: FetchOpts,
    session: SessionManager,
    report: IntegrityReport
): Promise<{ main: FormData; stageForms: FormData[] }> {
    const records: Record_[] = [];
    const columns = new Map<string, string>();
    const accept = makePageGuard(form.key);
    // stage id -> accumulated child rows / columns, across all pages.
    const stageRows = new Map<string, Record_[]>();
    const stageColumns = new Map<string, Map<string, string>>();
    const unresolvedHeaders = newUnresolvedHeaderTracker();
    let page = 1;

    for (;;) {
        const currentPage = page;
        const resp = await session.retryWithBackoff(() =>
            api.tracker.trackedEntities
                .get({
                    fields: {
                        trackedEntity: true,
                        orgUnit: true,
                        createdAt: true,
                        updatedAt: true,
                        attributes: { attribute: true, value: true },
                        enrollments: {
                            enrollment: true,
                            enrolledAt: true,
                            status: true,
                            events: {
                                event: true,
                                programStage: true,
                                occurredAt: true,
                                status: true,
                                dataValues: { dataElement: true, value: true },
                            },
                        },
                    },
                    program: form.uid,
                    ...ouParams(opts.orgUnit),
                    page: currentPage,
                    pageSize: opts.pageSize,
                    totalPages: false,
                })
                .getData()
        );

        const instances = resp.instances ?? [];
        if (!accept(instances.map(tei => tei.trackedEntity))) break;

        for (const tei of instances) {
            const enrollments = tei.enrollments ?? [];
            const first = enrollments[0];
            const { main, stageRecords } = buildTrackerRecords(
                {
                    trackedEntity: tei.trackedEntity,
                    orgUnit: tei.orgUnit,
                    createdAt: tei.createdAt,
                    updatedAt: tei.updatedAt,
                    attributes: tei.attributes,
                    enrollment: first && { enrolledAt: first.enrolledAt, status: first.status },
                    events: enrollments.flatMap(e => e.events ?? []),
                },
                form,
                meta,
                columns,
                stageColumns,
                report,
                unresolvedHeaders
            );
            records.push(main);
            for (const { stageId, record } of stageRecords) {
                const rows = stageRows.get(stageId);
                if (rows) rows.push(record);
                else stageRows.set(stageId, [record]);
            }
        }

        if (instances.length < opts.pageSize) break;
        page += 1;
    }
    console.log(`  ${form.key}: ${records.length} records`);

    // One batched rescue for every column that fell back to a raw id, applied before
    // stageForms are built below so patched labels flow into the child sheets too.
    await resolveStrayHeaderLabels(api, form.key, unresolvedHeaders, columns, stageColumns, session, report);

    // Sort by stage id (immutable, DHIS2-assigned) before assigning keys, so the result
    // is independent of stageRows' Map iteration order — which comes from API page/record
    // order and is not guaranteed stable run-to-run for a live, changing dataset. This is
    // the same order discovery's forecast uses, so predicted and actual sheet keys match.
    const usedKeys = new Set<string>([form.key]);
    const stageForms: FormData[] = [];
    const sortedStageEntries = [...stageRows.entries()].sort(([a], [b]) => a.localeCompare(b));
    for (const [stageId, rows] of sortedStageEntries) {
        const stage = meta.stageById.get(stageId);
        if (!stage) continue;

        const childKey = assignStageChildKey(form.key, stage, usedKeys);
        const childForm = buildStageChildForm(form, stage, childKey);
        stageForms.push({
            form: childForm,
            records: rows,
            columns: finalizeColumns(stageColumns.get(stageId) ?? new Map(), meta),
        });
    }

    return { main: { form, records, columns: finalizeColumns(columns, meta) }, stageForms };
}

async function fetchEvents(
    api: D2Api,
    form: ResolvedForm,
    meta: ProgramMeta,
    opts: FetchOpts,
    session: SessionManager,
    report: IntegrityReport
): Promise<FormData> {
    const records: Record_[] = [];
    const columns = new Map<string, string>();
    const accept = makePageGuard(form.key);
    const unresolvedDataElementIds = new Set<string>();
    let page = 1;

    for (;;) {
        const currentPage = page;
        const resp = await session.retryWithBackoff(() =>
            api.tracker.events
                .get({
                    fields: {
                        event: true,
                        orgUnit: true,
                        occurredAt: true,
                        createdAt: true,
                        updatedAt: true,
                        status: true,
                        dataValues: { dataElement: true, value: true },
                    },
                    program: form.uid,
                    ...ouParams(opts.orgUnit),
                    page: currentPage,
                    pageSize: opts.pageSize,
                    totalPages: false,
                })
                .getData()
        );

        const instances = resp.instances ?? [];
        if (!accept(instances.map(event => event.event))) break;

        for (const event of instances) {
            const values = new Map<string, string>();
            for (const dv of event.dataValues ?? []) {
                if (dv.value === undefined || dv.value === null) continue;
                values.set(dv.dataElement, String(dv.value));
                if (!columns.has(dv.dataElement)) {
                    const label = meta.dataElements.get(dv.dataElement);
                    columns.set(dv.dataElement, label ?? dv.dataElement);
                    if (!label) unresolvedDataElementIds.add(dv.dataElement);
                }
            }

            records.push({
                id: event.event,
                parentId: form.parentLinkField ? values.get(form.parentLinkField) ?? "" : "",
                // The Survey form is itself the root: its own id is the survey id.
                surveyId: form.surveyLinkField ? values.get(form.surveyLinkField) ?? "" : event.event,
                facilityId: "",
                orgUnit: event.orgUnit ?? "",
                label: values.get(PREVALENCE_SURVEY_NAME_DATAELEMENT_ID) ?? event.event,
                values,
                meta: new Map<string, string>([
                    ["occurred_at", event.occurredAt ?? ""],
                    ["created_at", event.createdAt ?? ""],
                    ["updated_at", event.updatedAt ?? ""],
                    ["status", event.status ?? ""],
                ]),
            });
        }

        if (instances.length < opts.pageSize) break;
        page += 1;
    }
    console.log(`  ${form.key}: ${records.length} records`);

    if (unresolvedDataElementIds.size > 0) {
        const ids = [...unresolvedDataElementIds];
        const rescued = await fetchDataElementLabels(api, ids, session);
        for (const id of ids) {
            const label = rescued.get(id);
            if (label) columns.set(id, label);
            else report.unresolvedHeaders.push({ formKey: form.key, id, kind: "dataElement" });
        }
        console.warn(
            `  ! ${form.key}: ${ids.length} column header(s) needed a direct-fetch rescue ` +
                `(program metadata didn't have the name)`
        );
    }

    return { form, records, columns: finalizeColumns(columns, meta) };
}

/**
 * Ward Summary Statistics is an aggregate dataSet, not a program: one row per
 * dataValue, keyed by orgUnit + period + attributeOptionCombo (the ward form id).
 */
async function fetchDataSet(api: D2Api, form: ResolvedForm, opts: FetchOpts): Promise<FormData> {
    if (!opts.orgUnit || !opts.startDate || !opts.endDate) {
        console.warn(
            `  ! ${form.key}: skipped (needs --org-unit, --start-date and --end-date; the ` +
                `dataValueSets endpoint requires an org unit and a period range).`
        );
        return { form, records: [], columns: new Map() };
    }

    const resp = await api.dataValues
        .getSet({
            dataSet: [form.uid],
            orgUnit: [opts.orgUnit],
            startDate: opts.startDate,
            endDate: opts.endDate,
            children: true,
        })
        .getData();

    const dataValues = resp.dataValues ?? [];
    const records: Record_[] = dataValues.map((dv, i: number) => ({
        id: `${dv.orgUnit}-${dv.period}-${dv.attributeOptionCombo}-${dv.dataElement}-${i}`,
        parentId: "",
        // Aggregate data: no survey FK. It links to Facility by org unit, and to a ward
        // event via attributeOptionCombo (= the ward form id).
        surveyId: "",
        facilityId: "",
        orgUnit: dv.orgUnit ?? "",
        label: "",
        values: new Map<string, string>([
            ["period", String(dv.period ?? "")],
            ["ward_form_id", String(dv.attributeOptionCombo ?? "")],
            ["data_element", String(dv.dataElement ?? "")],
            ["category_option_combo", String(dv.categoryOptionCombo ?? "")],
            ["value", String(dv.value ?? "")],
        ]),
        meta: new Map<string, string>([
            ["last_updated", dv.lastUpdated ?? ""],
            ["stored_by", dv.storedBy ?? ""],
        ]),
    }));

    const columns = new Map<string, string>([
        ["period", "Period"],
        ["ward_form_id", "Ward form id"],
        ["data_element", "Data element"],
        ["category_option_combo", "Category option combo"],
        ["value", "Value"],
    ]);

    console.log(`  ${form.key}: ${records.length} data values`);
    return { form, records, columns };
}

// --- Org unit names ------------------------------------------------------------

async function fetchOrgUnitNames(api: D2Api, ids: string[]): Promise<Map<string, string>> {
    const names = new Map<string, string>();
    const chunkSize = 200;

    for (let i = 0; i < ids.length; i += chunkSize) {
        const chunk = ids.slice(i, i + chunkSize);
        const { organisationUnits } = await api.metadata
            .get({
                organisationUnits: {
                    fields: { id: true, name: true },
                    filter: { id: { in: chunk } },
                },
            })
            .getData();
        for (const ou of organisationUnits) names.set(ou.id, ou.name);
    }

    return names;
}

// --- Workbook -----------------------------------------------------------------

const EXCEL_SHEET_NAME_MAX = 31;

/** A short, deterministic, order-independent id: a low-collision digest of the input,
 * NOT a cryptographic hash — this only needs to disambiguate a handful of sheet names. */
function shortStableId(input: string): string {
    let hash = 0x811c9dc5; // FNV-1a 32-bit offset basis
    for (let i = 0; i < input.length; i++) {
        hash ^= input.charCodeAt(i);
        hash = Math.imul(hash, 0x01000193) >>> 0; // FNV-1a prime, kept unsigned 32-bit
    }
    return hash.toString(36);
}

export type SheetNameSource = {
    /** Internal graph key (e.g. "Facility", "Facility__WardData"). Doubles as the label. */
    key: string;
    /** A stable, unique-per-source string to derive a disambiguation suffix from if the
     * sanitized/truncated name collides with another source's. Using each form's own DHIS2
     * uid (not an encounter-order counter) keeps the assignment stable across runs even if
     * API paging order changes. */
    stableId: string;
};

/**
 * Assigns every source a sheet name that is simultaneously: Excel-safe (forbidden chars
 * replaced, <=31 chars), collision-free against every other source AND against `reserved`,
 * and deterministic — processing order is a stable sort by `key`, so the result depends
 * only on the input names themselves, never on Map/array iteration order upstream.
 *
 * Any collision that needed a disambiguation suffix is appended to `report.sheetNameCollisions`
 * (if given) so it surfaces in `_index` rather than passing silently.
 */
export function assignSheetNames(
    sources: SheetNameSource[],
    reserved: string[] = [],
    report?: IntegrityReport
): Map<string, string> {
    const used = new Set<string>(reserved);
    const result = new Map<string, string>();

    for (const source of [...sources].sort((a, b) => a.key.localeCompare(b.key))) {
        const safeLabel = source.key.replace(/[:\\/?*[\]]/g, "-");
        let name = safeLabel.slice(0, EXCEL_SHEET_NAME_MAX);

        if (used.has(name)) {
            let attempt = 0;
            let candidate: string;
            do {
                const id = attempt === 0 ? shortStableId(source.stableId) : `${shortStableId(source.stableId)}${attempt}`;
                const suffix = `~${id}`;
                candidate = safeLabel.slice(0, Math.max(0, EXCEL_SHEET_NAME_MAX - suffix.length)) + suffix;
                attempt++;
            } while (used.has(candidate) && attempt < 1000);
            name = candidate;
            report?.sheetNameCollisions.push({ key: source.key, assignedName: name });
            console.warn(`  ! sheet name collision resolved: ${source.key} -> "${name}"`);
        }

        used.add(name);
        result.set(source.key, name);
    }

    return result;
}

export type IntegrityReport = {
    /** Records whose parentId does not resolve to a parent record. */
    orphans: { formKey: string; recordId: string; missingParentId: string }[];
    /** (survey, orgUnit) pairs matching more than one Facility record. */
    ambiguousFacilities: { surveyId: string; orgUnit: string; facilityIds: string[] }[];
    /** Records whose own surveyId disagrees with the survey reached via their parent. */
    surveyMismatches: { formKey: string; recordId: string; own: string; viaParent: string }[];
    /** Records for which no Facility could be derived. */
    unresolvedFacility: { formKey: string; count: number }[];
    /** Stage events whose programStage id did not resolve in program metadata (data was skipped). */
    unresolvedStageEvents: { formKey: string; programStage: string; count: number }[];
    /** Sheet names that needed a disambiguation suffix to stay unique. */
    sheetNameCollisions: { key: string; assignedName: string }[];
    /** Column headers that still show a raw DHIS2 id after every resolution tier ran
     * (stage config, the program's flat dataElement/attribute list, and a direct fetch). */
    unresolvedHeaders: { formKey: string; id: string; kind: "dataElement" | "trackedEntityAttribute" }[];
};

export function emptyIntegrityReport(): IntegrityReport {
    return {
        orphans: [],
        ambiguousFacilities: [],
        surveyMismatches: [],
        unresolvedFacility: [],
        unresolvedStageEvents: [],
        sheetNameCollisions: [],
        unresolvedHeaders: [],
    };
}

/** Records one occurrence of an event whose programStage id did not resolve in metadata. */
export function recordUnresolvedStageEvent(
    report: IntegrityReport,
    formKey: string,
    programStage: string
): void {
    const existing = report.unresolvedStageEvents.find(
        u => u.formKey === formKey && u.programStage === programStage
    );
    if (existing) existing.count += 1;
    else report.unresolvedStageEvents.push({ formKey, programStage, count: 1 });
}

export function reportIntegrity(report: IntegrityReport): void {
    const {
        orphans,
        ambiguousFacilities,
        surveyMismatches,
        unresolvedFacility,
        unresolvedStageEvents,
        sheetNameCollisions,
        unresolvedHeaders,
    } = report;
    if (
        orphans.length === 0 &&
        ambiguousFacilities.length === 0 &&
        surveyMismatches.length === 0 &&
        unresolvedFacility.length === 0 &&
        unresolvedStageEvents.length === 0 &&
        sheetNameCollisions.length === 0 &&
        unresolvedHeaders.length === 0
    ) {
        console.log(
            "  integrity: no orphans, no ambiguous facilities, no survey mismatches, no skipped events, no unresolved headers."
        );
        return;
    }

    console.warn("\n  ! Integrity findings (see the _index sheet for detail):");
    if (orphans.length > 0) {
        console.warn(`      orphans (parent id not found): ${orphans.length}`);
    }
    if (unresolvedFacility.length > 0) {
        for (const u of unresolvedFacility) {
            console.warn(`      ${u.formKey}: ${u.count} record(s) with no derivable Facility`);
        }
    }
    if (ambiguousFacilities.length > 0) {
        console.warn(
            `      ambiguous facilities ((survey, orgUnit) matching >1 Facility): ${ambiguousFacilities.length}`
        );
    }
    if (surveyMismatches.length > 0) {
        console.warn(
            `      survey mismatches (own Survey_id != Survey via parent): ${surveyMismatches.length}`
        );
    }
    if (unresolvedStageEvents.length > 0) {
        for (const u of unresolvedStageEvents) {
            console.warn(
                `      ${u.formKey}: ${u.count} event(s) on unresolvable stage ${u.programStage} SKIPPED (data lost)`
            );
        }
    }
    if (sheetNameCollisions.length > 0) {
        console.warn(`      sheet name collisions resolved: ${sheetNameCollisions.length}`);
    }
    if (unresolvedHeaders.length > 0) {
        console.warn(
            `      unresolved column headers (raw id shown, name not found anywhere): ${unresolvedHeaders.length}`
        );
        for (const h of unresolvedHeaders) {
            console.warn(`        ${h.formKey}: ${h.kind} ${h.id}`);
        }
    }
}

/**
 * Case report and the leaf forms have NO foreign key to Facility (see the note on
 * PARENT_BY_DEFAULT_UID). Facility is derived by matching on the pair
 * (Survey_id, org_unit_id) against the Facility-level form.
 *
 * Leaf forms inherit orgUnit from their Case report in the UI, but we match on each
 * record's own orgUnit, which is what the server actually stored.
 */
export function resolveFacilityIds(
    all: FormData[],
    facilityKey: string,
    report: IntegrityReport
): void {
    const facilityData = all.find(d => d.form.key === facilityKey);
    if (!facilityData) return;

    // (surveyId + orgUnit) -> facility record ids
    const index = new Map<string, string[]>();
    for (const facility of facilityData.records) {
        const composite = `${facility.surveyId}|${facility.orgUnit}`;
        const existing = index.get(composite);
        if (existing) existing.push(facility.id);
        else index.set(composite, [facility.id]);
    }

    for (const [composite, ids] of index) {
        if (ids.length > 1) {
            const [surveyId = "", orgUnit = ""] = composite.split("|");
            report.ambiguousFacilities.push({ surveyId, orgUnit, facilityIds: ids });
        }
    }

    for (const data of all) {
        if (data.form.key === facilityKey) {
            // A Facility record is its own facility.
            for (const record of data.records) record.facilityId = record.id;
            continue;
        }
        if (data.form.kind === "dataSet" || !data.form.parentKey) continue;

        let unresolved = 0;
        for (const record of data.records) {
            const matches = index.get(`${record.surveyId}|${record.orgUnit}`);
            // Ambiguous pairs are reported above; take the first for determinism.
            if (matches && matches[0]) record.facilityId = matches[0];
            else unresolved++;
        }
        if (unresolved > 0) report.unresolvedFacility.push({ formKey: data.form.key, count: unresolved });
    }
}

/**
 * Cross-check: a leaf's own Survey id (from its own TEA) against the Survey reached by
 * walking parent -> Case report -> Survey. Disagreement indicates a data integrity problem
 * server-side; we report it rather than silently trusting one side.
 */
export function crossCheckSurveyIds(
    all: FormData[],
    recordIndex: Map<string, Map<string, Record_>>,
    report: IntegrityReport
): void {
    for (const data of all) {
        if (!data.form.parentKey || data.form.kind === "dataSet") continue;

        for (const record of data.records) {
            if (!record.surveyId || !record.parentId) continue;
            const parent = recordIndex.get(data.form.parentKey)?.get(record.parentId);
            if (!parent) continue;

            const viaParent = parent.surveyId || parent.id;
            if (viaParent && record.surveyId !== viaParent) {
                report.surveyMismatches.push({
                    formKey: data.form.key,
                    recordId: record.id,
                    own: record.surveyId,
                    viaParent,
                });
            }
        }
    }
}

/**
 * Human-readable lineage, e.g. "Kenya PPS 2024 > Patient 0187 > D28". Walks the FK chain
 * (Survey -> Case report -> leaf); Facility is not on that chain because no FK exists, so
 * it does not appear here — see the Facility_id column instead.
 */
export function buildBreadcrumb(
    form: ResolvedForm,
    record: Record_,
    byKey: Map<string, FormData>,
    recordIndex: Map<string, Map<string, Record_>>,
    report?: IntegrityReport
): string {
    const parts: string[] = [];

    let currentForm: ResolvedForm | undefined = form;
    let currentRecord: Record_ | undefined = record;
    const seen = new Set<string>();

    while (currentForm?.parentKey) {
        const parentData = byKey.get(currentForm.parentKey);
        if (!parentData || !currentRecord) break;

        const parentId: string = currentRecord.parentId;
        if (!parentId || seen.has(parentId)) break;
        seen.add(parentId);

        const parentRecord: Record_ | undefined = recordIndex
            .get(currentForm.parentKey)
            ?.get(parentId);

        if (!parentRecord && report && currentRecord === record) {
            report.orphans.push({
                formKey: form.key,
                recordId: record.id,
                missingParentId: parentId,
            });
        }
        parts.unshift(parentRecord?.label || parentId);

        currentForm = parentData.form;
        currentRecord = parentRecord;
    }

    parts.push(record.label || record.id);
    return parts.join(" > ");
}

export type KeyColumn = { header: string; get: (record: Record_) => string };

/**
 * The ancestor key columns for a sheet, in a fixed order. Survey_id is read directly from
 * the record's own survey-link attribute (no chain walking); Facility_id is derived from
 * (Survey_id, org_unit_id); the immediate-parent column is only added when the parent is
 * not the Survey itself (i.e. CaseReport_id on the 8 leaf forms).
 */
export function keyColumnsFor(form: ResolvedForm, byKey: Map<string, FormData>): KeyColumn[] {
    if (form.kind === "dataSet") return [];

    const isSurvey = form.defaultUid === PREVALENCE_SURVEY_FORM_ID;
    const isFacility = form.defaultUid === PREVALENCE_FACILITY_LEVEL_FORM_ID;
    const cols: KeyColumn[] = [];

    if (!isSurvey) cols.push({ header: "Survey_id", get: r => r.surveyId });
    if (!isSurvey && !isFacility) cols.push({ header: "Facility_id", get: r => r.facilityId });

    const parentForm = form.parentKey ? byKey.get(form.parentKey)?.form : undefined;
    if (parentForm && parentForm.defaultUid !== PREVALENCE_SURVEY_FORM_ID) {
        const header = `${parentForm.key}_id`;
        // Avoid a duplicate header: e.g. a repeatable stage owned directly by Facility
        // would otherwise get "Facility_id" twice — once derived (isFacility check
        // above doesn't apply to the child), once as its immediate-parent key. Both
        // hold the same value in that case (the parent record IS the Facility record),
        // so keeping just the derived column is correct, not just cosmetic.
        if (!cols.some(c => c.header === header)) {
            cols.push({ header, get: r => r.parentId });
        }
    }
    return cols;
}

export function writeFormSheet(
    workbook: Excel.Workbook,
    data: FormData,
    byKey: Map<string, FormData>,
    recordIndex: Map<string, Map<string, Record_>>,
    orgUnitNames: Map<string, string>,
    sheetName: string,
    report?: IntegrityReport
): void {
    const { form, records, columns } = data;
    const sheet = workbook.addWorksheet(sheetName);

    const keyColumns = keyColumnsFor(form, byKey);
    const metaKeys = [...new Set(records.flatMap(r => [...r.meta.keys()]))];

    const header = [
        ...keyColumns.map(c => c.header),
        "path",
        "record_id",
        "org_unit_id",
        "org_unit_name",
        ...metaKeys,
        ...[...columns.values()],
    ];
    sheet.addRow(header);
    sheet.getRow(1).font = { bold: true };
    sheet.views = [{ state: "frozen", ySplit: 1 }];

    const columnUids = [...columns.keys()];

    if (records.length + 1 > EXCEL_MAX_ROWS) {
        console.warn(
            `  ! ${form.key}: ${records.length} rows exceeds Excel's sheet limit; only the first ` +
                `${EXCEL_MAX_ROWS - 1} are written. Narrow the scope (e.g. a smaller --org-unit) ` +
                `to extract the rest.`
        );
    }

    for (const record of records.slice(0, EXCEL_MAX_ROWS - 1)) {
        const breadcrumb = buildBreadcrumb(form, record, byKey, recordIndex, report);

        const row = [
            ...keyColumns.map(c => c.get(record)),
            breadcrumb,
            record.id,
            record.orgUnit,
            orgUnitNames.get(record.orgUnit) ?? "",
            ...metaKeys.map(k => record.meta.get(k) ?? ""),
            ...columnUids.map(uid => record.values.get(uid) ?? ""),
        ];
        sheet.addRow(row);
    }

    // Format as text at column level (one style record per column rather than
    // per cell) so DHIS2 UIDs and coded values are never re-parsed by Excel as
    // numbers or dates if the sheet is edited.
    sheet.columns.forEach(column => {
        column.width = 18;
        column.numFmt = "@";
    });
}

/**
 * Machine-readable relationship contract: one row per sheet giving its primary key, parent
 * key, ancestor keys, expected cardinality and an example join. Downstream tools (Power BI,
 * SQL loaders) can read this sheet instead of hard-coding the model.
 */
export function writeRelationshipsSheet(
    workbook: Excel.Workbook,
    all: FormData[],
    facilityKey: string | undefined
): void {
    const sheet = workbook.addWorksheet("_relationships");
    const byKey = new Map(all.map(d => [d.form.key, d]));

    sheet.addRow(["Relationship specification"]);
    sheet.getRow(1).font = { bold: true, size: 14 };
    sheet.addRow([
        "Note",
        "Case report has NO foreign key to Facility: it links to Survey. Facility_id is DERIVED " +
            "by matching (Survey_id, org_unit_id) against the Facility sheet.",
    ]);
    sheet.addRow([]);

    const header = sheet.addRow([
        "sheet",
        "primary_key",
        "parent_sheet",
        "parent_key",
        "ancestor_keys",
        "cardinality",
        "source_field",
        "example_join",
    ]);
    header.font = { bold: true };

    for (const { form } of all) {
        const keyCols = keyColumnsFor(form, byKey);
        const isSurvey = form.defaultUid === PREVALENCE_SURVEY_FORM_ID;
        const parentForm = form.parentKey ? byKey.get(form.parentKey)?.form : undefined;

        let parentKeyCol = "";
        let cardinality = "";
        let exampleJoin = "";
        let sourceField = form.parentLinkField || "";

        if (form.kind === "dataSet") {
            parentKeyCol = "(none)";
            cardinality = "aggregate: 1 row per dataElement x coc x period x orgUnit";
            exampleJoin = facilityKey
                ? `${form.key}.org_unit_id = ${facilityKey}.org_unit_id`
                : "org_unit_id = <facility org unit>";
            sourceField = "attributeOptionCombo = ward form id";
        } else if (isSurvey) {
            parentKeyCol = "(root)";
            cardinality = "root";
            exampleJoin = "-";
            sourceField = "-";
        } else if (parentForm) {
            const isFacility = form.defaultUid === PREVALENCE_FACILITY_LEVEL_FORM_ID;
            parentKeyCol = isFacility || parentForm.defaultUid === PREVALENCE_SURVEY_FORM_ID
                ? "Survey_id"
                : `${parentForm.key}_id`;
            cardinality = `${parentForm.key} 1 -> N ${form.key}`;
            exampleJoin = `${form.key}.${parentKeyCol} = ${parentForm.key}.record_id`;
        }

        sheet.addRow([
            form.key,
            "record_id",
            parentForm?.key ?? (form.kind === "dataSet" ? "(none)" : "(root)"),
            parentKeyCol,
            keyCols.map(c => c.header).join(", ") || "-",
            cardinality,
            sourceField || "-",
            exampleJoin,
        ]);
    }

    sheet.addRow([]);
    const derivedHeader = sheet.addRow(["Derived key", "Rule"]);
    derivedHeader.font = { bold: true };
    sheet.addRow([
        "Facility_id",
        facilityKey
            ? `= ${facilityKey}.record_id WHERE ${facilityKey}.Survey_id = <row>.Survey_id AND ${facilityKey}.org_unit_id = <row>.org_unit_id`
            : "(Facility sheet not extracted)",
    ]);
    sheet.addRow([
        "Survey_id",
        "Read directly from each record's own survey-link attribute — every record traces to its Survey without walking the chain.",
    ]);

    sheet.getColumn(1).width = 20;
    sheet.getColumn(2).width = 14;
    sheet.getColumn(3).width = 16;
    sheet.getColumn(4).width = 16;
    sheet.getColumn(5).width = 42;
    sheet.getColumn(6).width = 34;
    sheet.getColumn(7).width = 24;
    sheet.getColumn(8).width = 52;
}

export function writeIndexSheet(
    workbook: Excel.Workbook,
    all: FormData[],
    unresolved: { spec: FormSpec; reason: string }[],
    baseUrl: string,
    extra?: {
        country?: ResolvedCountry;
        integrity?: IntegrityReport;
        reconciliation?: { formKey: string; expected?: number; actual: number }[];
        discoveries?: FormDiscovery[];
    }
): void {
    const sheet = workbook.addWorksheet("_index");
    const byKey = new Map(all.map(d => [d.form.key, d]));

    sheet.addRow(["AMR Surveys extraction"]);
    sheet.getRow(1).font = { bold: true, size: 14 };
    sheet.addRow(["Server", baseUrl]);
    sheet.addRow(["Extracted at", new Date().toISOString()]);
    if (extra?.country) {
        const c = extra.country;
        sheet.addRow(["Country", `${c.name}${c.code ? ` (${c.code})` : ""} uid=${c.id} level=${c.level}`]);
    }
    sheet.addRow([
        "Note",
        "Read-only extraction. One sheet per form; join on the *_id columns. See _relationships.",
    ]);
    sheet.addRow([]);

    const headerRow = sheet.addRow([
        "Sheet",
        "Form name (server)",
        "UID",
        "Kind",
        "Custom",
        "Parent sheet",
        "Parent link field",
        "Records",
        "Expected",
        "Reconciled",
    ]);
    headerRow.font = { bold: true };

    const expectedByKey = new Map((extra?.reconciliation ?? []).map(r => [r.formKey, r.expected]));

    for (const { form, records } of all) {
        const expected = expectedByKey.get(form.key);
        const reconciled =
            expected === undefined ? "n/a" : expected === records.length ? "OK" : "MISMATCH";
        sheet.addRow([
            form.key,
            form.serverName,
            form.uid,
            form.kind,
            form.isCustom ? "yes" : "no",
            form.parentKey ?? "(root)",
            form.parentLinkField || "(none)",
            records.length,
            expected ?? "n/a",
            reconciled,
        ]);
    }

    sheet.addRow([]);
    const treeHeader = sheet.addRow(["Hierarchy (child counts)"]);
    treeHeader.font = { bold: true };

    const roots = all.filter(d => !d.form.parentKey);
    const renderTree = (data: FormData, depth: number) => {
        sheet.addRow([`${"    ".repeat(depth)}${depth > 0 ? "└─ " : ""}${data.form.key}`, data.records.length]);
        for (const child of all.filter(d => d.form.parentKey === data.form.key)) {
            renderTree(child, depth + 1);
        }
    };
    for (const root of roots) renderTree(root, 0);
    for (const orphan of all.filter(d => d.form.parentKey && !byKey.has(d.form.parentKey))) {
        renderTree(orphan, 0);
    }

    if (unresolved.length > 0) {
        sheet.addRow([]);
        const uHeader = sheet.addRow(["Not extracted", "Reason"]);
        uHeader.font = { bold: true };
        for (const { spec, reason } of unresolved) {
            sheet.addRow([spec.requestedName, reason]);
        }
    }

    const integrity = extra?.integrity;
    if (integrity) {
        sheet.addRow([]);
        const iHeader = sheet.addRow(["Integrity findings", "Detail"]);
        iHeader.font = { bold: true };

        const findings: [string, string][] = [];
        for (const o of integrity.orphans.slice(0, 200)) {
            findings.push([
                `orphan: ${o.formKey}`,
                `record ${o.recordId} references missing parent ${o.missingParentId}`,
            ]);
        }
        if (integrity.orphans.length > 200) {
            findings.push(["orphan: ...", `${integrity.orphans.length - 200} more not listed`]);
        }
        for (const u of integrity.unresolvedFacility) {
            findings.push([
                `no facility: ${u.formKey}`,
                `${u.count} record(s) had no (Survey_id, org_unit_id) match in the Facility sheet`,
            ]);
        }
        for (const a of integrity.ambiguousFacilities.slice(0, 50)) {
            findings.push([
                "ambiguous facility",
                `survey ${a.surveyId} + orgUnit ${a.orgUnit} matches ${a.facilityIds.length}: ${a.facilityIds.join(", ")}`,
            ]);
        }
        for (const m of integrity.surveyMismatches.slice(0, 50)) {
            findings.push([
                `survey mismatch: ${m.formKey}`,
                `record ${m.recordId} own Survey_id=${m.own} but via parent=${m.viaParent}`,
            ]);
        }
        for (const u of integrity.unresolvedStageEvents) {
            findings.push([
                `SKIPPED events: ${u.formKey}`,
                `${u.count} event(s) referenced unresolvable program stage ${u.programStage} — their data was NOT extracted`,
            ]);
        }
        for (const c of integrity.sheetNameCollisions) {
            findings.push([
                `sheet name collision: ${c.key}`,
                `assigned "${c.assignedName}" after a truncation/sanitization collision — no data lost, name changed`,
            ]);
        }
        for (const h of integrity.unresolvedHeaders) {
            findings.push([
                `unresolved header: ${h.formKey}`,
                `${h.kind} ${h.id} — its name was not found in the stage config, the program's ` +
                    `metadata, or a direct lookup; the column header shows the raw id`,
            ]);
        }

        if (findings.length === 0) sheet.addRow(["none", "No orphans, ambiguous facilities or survey mismatches."]);
        else for (const f of findings) sheet.addRow(f);
    }

    sheet.getColumn(1).width = 34;
    sheet.getColumn(2).width = 62;
    sheet.getColumn(3).width = 14;
    for (let i = 4; i <= 10; i++) sheet.getColumn(i).width = 14;
}

// --- Country resolution & discovery -------------------------------------------

export type ResolvedCountry = {
    id: string;
    name: string;
    code?: string;
    level: number;
    matchedBy: string;
};

const UID_RE = /^[A-Za-z][A-Za-z0-9]{10}$/;

/**
 * Resolves --country by code (e.g. KEN), then exact name, then UID. Fails loudly and lists
 * candidates on ambiguity rather than silently picking one.
 */
export async function resolveCountry(api: D2Api, input: string): Promise<ResolvedCountry> {
    const term = input.trim();
    const fields = { id: true, name: true, code: true, level: true } as const;

    const attempts = [
        { matchedBy: "code", filter: { code: { eq: term.toUpperCase() } } },
        { matchedBy: "name", filter: { name: { eq: term } } },
        ...(UID_RE.test(term) ? [{ matchedBy: "uid", filter: { id: { eq: term } } }] : []),
    ];

    for (const attempt of attempts) {
        const { organisationUnits } = await api.metadata
            .get({ organisationUnits: { fields, filter: attempt.filter } })
            .getData();

        if (organisationUnits.length === 1) {
            const ou = organisationUnits[0]!;
            return { ...ou, matchedBy: attempt.matchedBy };
        }
        if (organisationUnits.length > 1) {
            throw new Error(
                `--country "${term}" is ambiguous by ${attempt.matchedBy}: ` +
                    organisationUnits.map(o => `${o.name} (${o.id})`).join("; ")
            );
        }
    }
    throw new Error(
        `Could not resolve --country "${term}" by code, name or UID. ` +
            `Try the ISO3 code (e.g. KEN), the exact org unit name, or its UID.`
    );
}

export type StageDiscovery = {
    /** The exact sheet key extraction would assign — computed via the same shared
     * assignStageChildKey()/buildStageChildForm() helpers fetchTracker uses, so this is a
     * prediction that cannot drift from what a real run would produce. */
    sheetKey: string;
    stageName: string;
    parentFormKey: string;
    /** Row count for this repeatable stage under the scope (probed server-side). */
    rows: number;
    /** Column count this stage's own sheet would have (ancestor keys + fixed + stage fields). */
    columns: number;
};

export type FormDiscovery = {
    form: ResolvedForm;
    /** Record count for this form under the country, or undefined if not probeable. */
    total?: number;
    minDate?: string;
    maxDate?: string;
    /** Expected column count on this form's own (main) sheet, when staticly determinable. */
    columns?: number;
    /** Repeatable-stage child sheets this form would produce. */
    stages: StageDiscovery[];
};

/** Column count for a form's own sheet: ancestor keys + fixed columns + meta + value columns. */
export function computeColumnCount(
    form: ResolvedForm,
    byKey: Map<string, FormData>,
    metaKeyCount: number,
    valueColumnCount: number
): number {
    return keyColumnsFor(form, byKey).length + FIXED_NON_META_COLUMNS + metaKeyCount + valueColumnCount;
}

/**
 * Probes one repeatable stage's row count directly — DHIS2's events endpoint accepts a
 * `programStage` filter, so this is one more pageSize=1/totalPages=true call, exactly as
 * cheap as the top-level probes above. Not a guess: an accurate server-side count.
 */
async function probeStageRows(
    api: D2Api,
    programUid: string,
    stageId: string,
    orgUnit: string,
    session: SessionManager
): Promise<number> {
    const resp = await session.retryWithBackoff(() =>
        api.tracker.events
            .get({
                fields: { event: true },
                program: programUid,
                programStage: stageId,
                orgUnit,
                ouMode: "DESCENDANTS",
                page: 1,
                pageSize: 1,
                totalPages: true,
            })
            .getData()
    );
    return resp.total ?? 0;
}

/**
 * Cheap preflight probe: for a tracker/event form, 1 count call + 2 date calls (pageSize=1),
 * plus 1 metadata call (needed to know column counts and enumerate repeatable stages), plus
 * 1 row-count call per repeatable stage. Deliberately per-FORM (and per-STAGE), never
 * per-facility — probing each facility would be O(forms x facilities) calls for no extra
 * information the country-wide scope doesn't already give us.
 */
async function discoverForm(
    api: D2Api,
    form: ResolvedForm,
    orgUnit: string,
    session: SessionManager,
    byKey: Map<string, FormData>
): Promise<FormDiscovery> {
    if (form.kind === "dataSet") {
        // dataValueSets has no cheap count/period probe; scope comes from --start-date/--end-date.
        return { form, stages: [] };
    }

    if (form.kind === "event") {
        const probe = (direction: "asc" | "desc") =>
            session.retryWithBackoff(() =>
                api.tracker.events
                    .get({
                        fields: { event: true, occurredAt: true },
                        program: form.uid,
                        orgUnit,
                        ouMode: "DESCENDANTS",
                        page: 1,
                        pageSize: 1,
                        totalPages: true,
                        order: `occurredAt:${direction}`,
                    })
                    .getData()
            );

        const [first, last, meta] = await Promise.all([
            probe("asc"),
            probe("desc"),
            fetchProgramMeta(api, form.uid),
        ]);
        return {
            form,
            total: first.total ?? 0,
            minDate: first.instances[0]?.occurredAt,
            maxDate: last.instances[0]?.occurredAt,
            columns: computeColumnCount(form, byKey, EVENT_META_KEYS.length, meta.dataElements.size),
            stages: [],
        };
    }

    // tracker
    const probe = (direction: "asc" | "desc") =>
        session.retryWithBackoff(() =>
            api.tracker.trackedEntities
                .get({
                    fields: { trackedEntity: true, createdAt: true },
                    program: form.uid,
                    orgUnit,
                    ouMode: "DESCENDANTS",
                    page: 1,
                    pageSize: 1,
                    totalPages: true,
                    order: [{ type: "field", field: "createdAt", direction }],
                })
                .getData()
        );

    const [first, last, meta] = await Promise.all([probe("asc"), probe("desc"), fetchProgramMeta(api, form.uid)]);

    const nonRepeatableDataElementCount = [...meta.stageById.values()]
        .filter(s => !s.repeatable)
        .reduce((sum, s) => sum + s.dataElementLabels.size, 0);
    const mainValueColumns = meta.attributes.size + nonRepeatableDataElementCount;

    const repeatableStages = [...meta.stageById.values()]
        .filter(s => s.repeatable)
        .sort((a, b) => a.id.localeCompare(b.id));
    const stageRowCounts = await mapWithConcurrency(repeatableStages, FETCH_CONCURRENCY, s =>
        probeStageRows(api, form.uid, s.id, orgUnit, session)
    );

    const usedKeys = new Set<string>([form.key]);
    const stages: StageDiscovery[] = repeatableStages.map((stage, idx) => {
        const sheetKey = assignStageChildKey(form.key, stage, usedKeys);
        const childForm = buildStageChildForm(form, stage, sheetKey);
        return {
            sheetKey,
            stageName: stage.name,
            parentFormKey: form.key,
            rows: stageRowCounts[idx] ?? 0,
            columns: computeColumnCount(
                childForm,
                byKey,
                TRACKER_STAGE_CHILD_META_KEYS.length,
                stage.dataElementLabels.size
            ),
        };
    });

    return {
        form,
        total: first.total ?? 0,
        minDate: first.instances[0]?.createdAt,
        maxDate: last.instances[0]?.createdAt,
        columns: computeColumnCount(form, byKey, TRACKER_MAIN_META_KEYS.length, mainValueColumns),
        stages,
    };
}

/** Runs the probes with bounded concurrency (mirrors glass-dev's FETCH_CONCURRENCY). */
export async function discover(
    api: D2Api,
    forms: ResolvedForm[],
    orgUnit: string,
    session: SessionManager
): Promise<FormDiscovery[]> {
    // A minimal, records-free FormData map so discoverForm can reuse the exact same
    // keyColumnsFor() the real writer uses — guaranteeing predicted and actual column
    // sets can never diverge.
    const byKey = new Map<string, FormData>(
        forms.map(f => [f.key, { form: f, records: [], columns: new Map() }])
    );

    return mapWithConcurrency(forms, FETCH_CONCURRENCY, f => discoverForm(api, f, orgUnit, session, byKey));
}

function shortDate(iso: string | undefined): string {
    return iso ? iso.slice(0, 10) : "-";
}

// --- Workbook-size forecast -----------------------------------------------------

export type SheetForecast = {
    sheetKey: string;
    parentFormKey?: string;
    /** Exact: probed server-side. */
    rows?: number;
    /**
     * UPPER BOUND, not a prediction. Counts every field CONFIGURED on the form, but a
     * sheet only gets a column for a field at least one record actually fills in — most
     * configured fields are typically empty. Measured on KGZ: a stage configured with 222
     * fields produced 23 real columns, and the whole workbook came out ~12 MB against a
     * 63-190 MB bound. Treat the size estimate as "definitely no bigger than".
     */
    columns?: number;
    cells?: number;
};

export type WorkbookForecast = {
    sheets: SheetForecast[];
    /** UPPER BOUND, not a prediction — see the note on `columns` in SheetForecast. */
    totalCells: number;
    /** Upper-bound compressed .xlsx size range, derived from totalCells. */
    estimatedSizeBytesRange: [number, number];
    warnings: string[];
};

// Rough, clearly-labeled compressed-xlsx bytes/cell range, informed by an earlier live
// comparison against this same server (see README). Deliberately wide, not false-precise.
const BYTES_PER_CELL_LOW = 15;
const BYTES_PER_CELL_HIGH = 45;
const LARGE_SHEET_COLUMN_WARNING = 200;
const LARGE_SHEET_ROW_WARNING = 900_000; // 90% of Excel's 1,048,576-row cap
const LARGE_WORKBOOK_CELL_WARNING = 2_000_000;

export function formatBytes(bytes: number): string {
    if (bytes < 1024) return `${Math.round(bytes)} B`;
    const mb = bytes / (1024 * 1024);
    return mb < 1024 ? `${mb.toFixed(1)} MB` : `${(mb / 1024).toFixed(2)} GB`;
}

/**
 * Turns discovery results into a per-sheet row/column/cell forecast, a total workbook
 * size estimate, and warnings for unusually large sheets — everything needed to describe
 * the workbook BEFORE extraction runs (instruction #4).
 *
 * Mirrors extract()'s own skip rule exactly (`total !== 0`, i.e. only a PROVEN-empty form
 * is skipped) so this never predicts a sheet a real run wouldn't produce, or vice versa —
 * that equivalence is what makes "discovery accurately predicts the export" true rather
 * than aspirational.
 */
export function forecastWorkbook(discoveries: FormDiscovery[]): WorkbookForecast {
    const sheets: SheetForecast[] = [];
    const warnings: string[] = [];

    const addSheet = (f: SheetForecast) => {
        sheets.push(f);
        if (f.columns !== undefined && f.columns > LARGE_SHEET_COLUMN_WARNING) {
            warnings.push(`${f.sheetKey}: ${f.columns} columns — wide sheet, slow to browse by hand in Excel`);
        }
        if (f.rows !== undefined && f.rows > LARGE_SHEET_ROW_WARNING) {
            warnings.push(`${f.sheetKey}: ${f.rows.toLocaleString()} rows — approaching Excel's 1,048,576-row limit`);
        }
    };

    for (const d of discoveries.filter(d => d.total !== 0)) {
        addSheet({
            sheetKey: d.form.key,
            rows: d.total,
            columns: d.columns,
            cells: d.total !== undefined && d.columns !== undefined ? d.total * d.columns : undefined,
        });
        // A stage with a proven-zero row count produces no sheet at extraction time
        // either (fetchTracker only creates one once at least one event exists).
        for (const s of d.stages.filter(s => s.rows !== 0)) {
            addSheet({
                sheetKey: s.sheetKey,
                parentFormKey: s.parentFormKey,
                rows: s.rows,
                columns: s.columns,
                cells: s.rows * s.columns,
            });
        }
    }

    const totalCells = sheets.reduce((sum, s) => sum + (s.cells ?? 0), 0);
    if (totalCells > LARGE_WORKBOOK_CELL_WARNING) {
        warnings.push(
            `workbook total: ~${totalCells.toLocaleString()} cells — large file; extraction and opening it in Excel may be slow`
        );
    }

    return {
        sheets,
        totalCells,
        estimatedSizeBytesRange: [totalCells * BYTES_PER_CELL_LOW, totalCells * BYTES_PER_CELL_HIGH],
        warnings,
    };
}

export function reportDiscovery(
    country: ResolvedCountry | undefined,
    rootOrgUnit: string,
    discoveries: FormDiscovery[],
    unresolved: { spec: FormSpec; reason: string }[]
): void {
    if (country) {
        console.log(
            `\nCountry resolved: ${country.name}` +
                `${country.code ? ` (${country.code})` : ""} uid=${country.id} ` +
                `level=${country.level} [matched by ${country.matchedBy}]`
        );
    } else {
        console.log(`\nScope: org unit ${rootOrgUnit} (with descendants)`);
    }

    console.log("\nForms detected:");
    console.log(
        "  " +
            [
                "SHEET".padEnd(18),
                "KIND".padEnd(8),
                "CUSTOM".padEnd(7),
                "RECORDS".padStart(9),
                "COLS".padStart(6),
                "  FIRST".padEnd(13),
                "LAST",
            ].join("")
    );
    for (const d of discoveries) {
        const total = d.total === undefined ? "n/a" : String(d.total);
        const cols = d.columns === undefined ? "n/a" : String(d.columns);
        console.log(
            "  " +
                [
                    d.form.key.padEnd(18),
                    d.form.kind.padEnd(8),
                    (d.form.isCustom ? "yes" : "no").padEnd(7),
                    total.padStart(9),
                    cols.padStart(6),
                    ("  " + shortDate(d.minDate)).padEnd(13),
                    shortDate(d.maxDate),
                ].join("")
        );
        for (const s of d.stages) {
            // Child sheet keys are inherently longer than top-level ones (Owner__StageName),
            // so a fixed-width column would overrun and mangle the row (verified live: e.g.
            // "Facility__WardData" already exceeds the 18-char budget above). A free-form,
            // explicitly-labelled line stays readable at any key length instead.
            console.log(
                `      ${s.sheetKey}` +
                    `  rows=${s.rows}  columns=${s.columns}` +
                    `  (repeatable stage "${s.stageName}" of ${s.parentFormKey})`
            );
        }
    }

    const withData = discoveries.filter(d => (d.total ?? 0) > 0);
    const empty = discoveries.filter(d => d.total === 0);
    const totalRecords = discoveries.reduce((sum, d) => sum + (d.total ?? 0), 0);
    const totalStageRecords = discoveries.reduce(
        (sum, d) => sum + d.stages.reduce((s, x) => s + (x.rows ?? 0), 0),
        0
    );

    console.log(`\nPlanned extraction scope:`);
    console.log(`  forms with data      : ${withData.length}/${discoveries.length}`);
    console.log(`  main-sheet records   : ${totalRecords}`);
    if (totalStageRecords > 0) {
        console.log(`  repeatable-stage rows: ${totalStageRecords}`);
    }
    if (empty.length > 0) {
        console.log(`  skipped (0 records): ${empty.map(d => d.form.key).join(", ")}`);
    }
    const dataSets = discoveries.filter(d => d.form.kind === "dataSet");
    if (dataSets.length > 0) {
        console.log(
            `  ${dataSets.map(d => d.form.key).join(", ")}: aggregate dataSet — scope set by --start-date/--end-date`
        );
    }
    if (unresolved.length > 0) {
        console.log(`  not extracted: ${unresolved.map(u => u.spec.requestedName).join("; ")}`);
    }

    const forecast = forecastWorkbook(discoveries);
    console.log(`\nWorkbook forecast:`);
    console.log(`  worksheets    : ${2 + forecast.sheets.length}  (_index, _relationships + ${forecast.sheets.length} form sheet(s))`);
    console.log(`  rows          : exact (probed)`);
    console.log(`  columns/cells : UPPER BOUND — every configured field is counted, but a sheet`);
    console.log(`                  only gets a column for a field some record actually fills in.`);
    console.log(`  max cells     : ~${forecast.totalCells.toLocaleString()}`);
    console.log(
        `  max size      : ${formatBytes(forecast.estimatedSizeBytesRange[0])} – ${formatBytes(
            forecast.estimatedSizeBytesRange[1]
        )}  (real files are typically several times smaller)`
    );
    if (forecast.warnings.length > 0) {
        console.log(`  warnings (based on the upper bound):`);
        for (const w of forecast.warnings) console.log(`    ! ${w}`);
    } else {
        console.log(`  no size/shape warnings`);
    }
}

/** Facilities that actually hold data, derived from the extracted Facility records (no extra calls). */
export function facilitiesWithData(all: FormData[], facilityKey: string): Set<string> {
    const facilityData = all.find(d => d.form.key === facilityKey);
    return new Set((facilityData?.records ?? []).map(r => r.orgUnit).filter(Boolean));
}

// --- Main ---------------------------------------------------------------------

async function extract(args: {
    output?: string;
    country?: string;
    orgUnit?: string;
    startDate?: string;
    endDate?: string;
    pageSize: number;
    dryRun: boolean;
    discover: boolean;
    forms?: string;
}) {
    const envVars = getEnvVars();
    const baseUrl = envVars.url.replace(/\/+$/, "");
    const envLabel = deriveEnvLabel(baseUrl);

    console.log(`Connecting to ${baseUrl} (${describeAuth(envVars)}) [env: ${envLabel}]`);
    const api = getD2APiFromInstance(getInstance(envVars));

    // Forces DHIS2 to fully initialize the session before any real call (PAT bug workaround).
    await warmUpSession(api);
    const session = createSessionManager(api);

    const info = await api.system.info.getData();
    console.log(`  DHIS2 ${info.version}`);

    const modules = await fetchModules(api);
    console.log(`  datastore modules: ${modules.map(m => m.name).join(", ") || "(none)"}`);

    const wanted = args.forms
        ? FORMS.filter(f => args.forms!.split(",").map(s => s.trim()).includes(f.key))
        : FORMS;
    if (wanted.length === 0) throw new Error(`No forms matched --forms=${args.forms}`);

    const { resolved, unresolved } = await resolveForms(api, modules, wanted);
    linkParents(resolved);
    reportResolution(resolved, unresolved);

    if (args.dryRun) {
        console.log("\nDry run: no data extracted, no file written.");
        return;
    }
    if (resolved.length === 0) throw new Error("No forms could be resolved; nothing to extract.");

    // --- Discovery ------------------------------------------------------------
    // Scope: an explicit --org-unit wins; otherwise --country; otherwise everything readable.
    const country = args.country ? await resolveCountry(api, args.country) : undefined;
    const rootOrgUnit = args.orgUnit ?? country?.id;

    let discoveries: FormDiscovery[] = [];
    if (rootOrgUnit) {
        console.log("\nDiscovering ...");
        discoveries = await discover(api, resolved, rootOrgUnit, session);
        reportDiscovery(country, rootOrgUnit, discoveries, unresolved);
    } else if (args.discover) {
        throw new Error("--discover requires --country (or --org-unit) to scope the probes.");
    }

    if (args.discover) {
        console.log("\nDiscovery only: no data extracted, no file written.");
        return;
    }

    // Only skip forms proven to have exactly zero records — never narrow on anything else.
    const totalByKey = new Map(discoveries.map(d => [d.form.key, d.total]));
    const toExtract = resolved.filter(f => totalByKey.get(f.key) !== 0);
    const skipped = resolved.filter(f => totalByKey.get(f.key) === 0);
    if (skipped.length > 0) {
        console.log(`\nSkipping ${skipped.length} form(s) with 0 records: ${skipped.map(f => f.key).join(", ")}`);
    }

    console.log("\nExtracting ...");
    const opts: FetchOpts = {
        orgUnit: rootOrgUnit,
        pageSize: args.pageSize,
        startDate: args.startDate,
        endDate: args.endDate,
    };

    // Declared before extraction (not after) so fetchTracker can record diagnostics
    // (e.g. unresolved stage events) as they happen, not as an afterthought.
    const integrity = emptyIntegrityReport();

    // Forms are independent of each other, so they are fetched with bounded concurrency
    // rather than one at a time — this is the dominant cost of a run. `mapWithConcurrency`
    // preserves input order, so sheet order stays deterministic regardless of which form
    // happens to finish first.
    const perForm = await mapWithConcurrency(toExtract, EXTRACT_CONCURRENCY, async form => {
        if (form.kind === "dataSet") {
            return [await session.retryWithBackoff(() => fetchDataSet(api, form, opts))];
        }
        const meta = await session.retryWithBackoff(() => fetchProgramMeta(api, form.uid));
        if (form.kind !== "tracker") {
            return [await fetchEvents(api, form, meta, opts, session, integrity)];
        }
        const { main, stageForms } = await fetchTracker(api, form, meta, opts, session, integrity);
        for (const stage of stageForms) {
            console.log(`  ${stage.form.key}: ${stage.records.length} rows (repeatable stage)`);
        }
        return [main, ...stageForms];
    });
    const all: FormData[] = perForm.flat();

    const byKey = new Map(all.map(d => [d.form.key, d]));
    const recordIndex = new Map(
        all.map(d => [d.form.key, new Map(d.records.map(r => [r.id, r]))])
    );

    // --- Integrity ------------------------------------------------------------
    const facilityForm = resolved.find(f => f.defaultUid === PREVALENCE_FACILITY_LEVEL_FORM_ID);
    if (facilityForm) resolveFacilityIds(all, facilityForm.key, integrity);
    crossCheckSurveyIds(all, recordIndex, integrity);

    // Reconciliation: what discovery said exists vs what we actually pulled. This is what
    // proves the skip/scope optimisation above never silently dropped data.
    const reconciliation = all.map(d => ({
        formKey: d.form.key,
        expected: totalByKey.get(d.form.key),
        actual: d.records.length,
    }));
    const mismatches = reconciliation.filter(r => r.expected !== undefined && r.expected !== r.actual);
    if (mismatches.length > 0) {
        console.warn("\n  ! RECONCILIATION MISMATCH — extracted count != discovered count:");
        for (const m of mismatches) {
            console.warn(`      ${m.formKey}: expected ${m.expected}, got ${m.actual}`);
        }
    } else if (discoveries.length > 0) {
        console.log("\nReconciliation: all forms match their discovered counts.");
    }

    const orgUnitIds = [...new Set(all.flatMap(d => d.records.map(r => r.orgUnit)).filter(Boolean))];
    console.log(`\nResolving ${orgUnitIds.length} org unit names ...`);
    const orgUnitNames = await fetchOrgUnitNames(api, orgUnitIds);

    console.log("Writing workbook ...");
    const workbook = new Excel.Workbook();
    workbook.creator = "amr-surveys extract-forms";
    workbook.created = new Date();

    // Collision-safe, deterministic Excel tab names for every sheet, computed once up
    // front so a truncation/sanitization collision can never silently overwrite another
    // sheet. "_index" and "_relationships" are reserved so a (bizarre) form name can
    // never collide with them either.
    const sheetNames = assignSheetNames(
        all.map(d => ({ key: d.form.key, stableId: d.form.uid })),
        ["_index", "_relationships"],
        integrity
    );

    writeIndexSheet(workbook, all, unresolved, baseUrl, {
        country,
        integrity,
        reconciliation,
        discoveries,
    });
    writeRelationshipsSheet(workbook, all, facilityForm?.key);
    for (const data of all) {
        const name = sheetNames.get(data.form.key) ?? data.form.key.slice(0, EXCEL_SHEET_NAME_MAX);
        writeFormSheet(workbook, data, byKey, recordIndex, orgUnitNames, name, integrity);
    }

    // Name the file after the instance and country so an extract is never ambiguous
    // about where it came from (mirrors glass-dev's deriveEnvLabel convention).
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    const countryPart = country ? `${(country.code || country.name).toLowerCase()}_` : "";
    const defaultOut = `extracts/amr-surveys_${envLabel}_${countryPart}${timestamp}.xlsx`;

    const outPath = path.resolve(args.output ?? defaultOut);
    fs.mkdirSync(path.dirname(outPath), { recursive: true });
    await workbook.xlsx.writeFile(outPath);

    const total = all.reduce((sum, d) => sum + d.records.length, 0);
    console.log(`\nDone. ${total} records across ${all.length} sheets -> ${outPath}`);

    const facilities = facilityForm ? facilitiesWithData(all, facilityForm.key) : new Set();
    if (facilities.size > 0) console.log(`  facilities with data: ${facilities.size}`);
    reportIntegrity(integrity);
}

function main() {
    // Env is loaded by `-r dotenv/config` (see package.json), with
    // DOTENV_CONFIG_PATH=.env.local — same convention as the glass-dev scripts.
    const cmd = command({
        name: "extract-forms",
        description: "Read-only extraction of AMR Surveys forms into an Excel workbook",
        args: {
            output: option({
                type: optional(string),
                long: "output",
                short: "o",
                description:
                    "Output .xlsx path. Defaults to extracts/amr-surveys_<env>_<country>_<timestamp>.xlsx",
            }),
            country: option({
                type: optional(string),
                long: "country",
                description:
                    "Country to extract: ISO3 code (KEN), exact org unit name (Kenya), or UID. Scopes everything to that subtree.",
            }),
            orgUnit: option({
                type: optional(string),
                long: "org-unit",
                description:
                    "Root org unit UID (with descendants). Overrides --country. Omit both to extract everything the user can read.",
            }),
            startDate: option({
                type: optional(string),
                long: "start-date",
                description: "Start date YYYY-MM-DD (Ward Summary Statistics only)",
            }),
            endDate: option({
                type: optional(string),
                long: "end-date",
                description: "End date YYYY-MM-DD (Ward Summary Statistics only)",
            }),
            pageSize: option({
                type: number,
                long: "page-size",
                defaultValue: () => 500,
                description: "Records per API request",
            }),
            forms: option({
                type: optional(string),
                long: "forms",
                description: "Comma-separated subset of form keys (default: all)",
            }),
            dryRun: flag({
                type: boolean,
                long: "dry-run",
                description: "Resolve names to UIDs and print the mapping without extracting",
            }),
            discover: flag({
                type: boolean,
                long: "discover",
                description:
                    "Preflight only: report country, forms detected, record counts, min/max dates and the planned scope. Extracts nothing.",
            }),
        },
        handler: async args => {
            try {
                await extract(args);
            } catch (err) {
                console.error(`\nFailed: ${err instanceof Error ? err.message : String(err)}`);
                process.exit(1);
            }
        },
    });

    run(cmd, process.argv.slice(2));
}

// Only run when invoked directly, so the pure helpers above stay importable by tests.
if (require.main === module) {
    main();
}
