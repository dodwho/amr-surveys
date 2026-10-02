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
 * A form is defined by its DEFAULT program. A survey may swap in a custom variant (datastore
 * amr-surveys/modules .customForms), so a country's data lives in whichever of the default
 * and custom programs it used. Every candidate program is probed under the requested scope;
 * those with data are extracted and stacked into one sheet per form (see mergeByFormKey).
 *
 * Usage:
 *   yarn extract-forms --dry-run
 *   yarn extract-forms --country KEN                 one workbook for one country
 *   yarn extract-forms --per-country                 one workbook per country that has data
 *   yarn extract-forms                               everything the user can read, one workbook
 *   yarn extract-forms --country KEN --start-date 2024-01-01 --end-date 2024-12-31
 *
 * --start-date/--end-date only affect the Ward Summary Statistics dataSet (default: all time).
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
    WARD_STATISTICS_WARD_LEVEL_FORM_ID,
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
    /** Name as it appears on the DHIS2 server. Display only; resolution is by UID. */
    requestedName: string;
    /**
     * UID of the default program (or dataSet) from D2Survey.ts. The form is identified by
     * this; its custom variants are read from the datastore (see candidateProgramIds).
     */
    defaultUid: string;
    /** Key of the parent form in this registry, if any. */
    parentKey?: string;
    /** Forced kind; otherwise derived from server metadata. */
    kind?: FormKind;
};

const FORMS: FormSpec[] = [
    { key: "Survey", requestedName: "Survey form", defaultUid: PREVALENCE_SURVEY_FORM_ID },
    {
        key: "Facility",
        requestedName: "Facility-level form",
        defaultUid: PREVALENCE_FACILITY_LEVEL_FORM_ID,
        parentKey: "Survey",
    },
    {
        // Parent is Survey, NOT Facility — see the note on PARENT_BY_DEFAULT_UID.
        key: "CaseReport",
        requestedName: "Case report form - custom v1",
        defaultUid: PREVALENCE_CASE_REPORT_FORM_ID,
        parentKey: "Survey",
    },
    {
        key: "SampleShipment",
        requestedName: "Sample shipment and tracking form - custom v1",
        defaultUid: PREVALENCE_SAMPLE_SHIP_TRACK_FORM_ID,
        parentKey: "CaseReport",
    },
    {
        key: "CentralRefLab",
        requestedName: "Central reference laboratory ID/AST results form",
        defaultUid: PREVALENCE_CENTRAL_REF_LAB_FORM_ID,
        parentKey: "CaseReport",
    },
    {
        key: "PathogenIsolates",
        requestedName: "Pathogen Isolates storage and tracking log",
        defaultUid: PREVALENCE_PATHOGEN_ISO_STORE_TRACK_ID,
        parentKey: "CaseReport",
    },
    {
        key: "Supranational",
        requestedName: "Supranational Reference Laboratory ID/AST results form",
        defaultUid: PREVALENCE_SUPRANATIONAL_REF_LAB_ID,
        parentKey: "CaseReport",
    },
    {
        key: "FollowUpD28",
        requestedName: "Follow-up form D28 - custom v1",
        defaultUid: PREVALENCE_MORTALITY_FOLLOWUP_FORM,
        parentKey: "CaseReport",
    },
    {
        key: "DischargeClinical",
        requestedName: "Discharge form - Clinical Evaluation",
        defaultUid: PREVALENCE_MORTALITY_DISCHARGE_CLINICAL_FORM,
        parentKey: "CaseReport",
    },
    {
        key: "DischargeEconomic",
        requestedName: "Discharge form - Economical Evaluation",
        defaultUid: PREVALENCE_MORTALITY_DISCHARGE_ECONOMIC_FORM,
        parentKey: "CaseReport",
    },
    {
        key: "CohortEnrolment",
        requestedName: "Cohort 3 enrolment form - custom v1",
        defaultUid: PREVALENCE_MORTALITY_COHORT_ENORL_FORM,
        parentKey: "CaseReport",
    },
    {
        // Ward x specialty. Its successor below holds ward-level values; both are live.
        key: "WardSummaryStats",
        requestedName: "Ward Summary Statistics",
        defaultUid: WARD_SUMMARY_STATISTICS_FORM_ID,
        kind: "dataSet",
    },
    {
        key: "WardLevelStats",
        requestedName: "Ward Statistics for Ward",
        defaultUid: WARD_STATISTICS_WARD_LEVEL_FORM_ID,
        kind: "dataSet",
    },
];

// --- Types --------------------------------------------------------------------

/**
 * ONE program (or dataSet) backing a form. A form with a custom variant resolves to several
 * of these sharing a `key`: the default program plus each custom program.
 */
export type ResolvedForm = FormSpec & {
    uid: string;
    serverName: string;
    kind: FormKind;
    /** True when `uid` is a custom variant of the default program, not the default itself. */
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
    /** Data-quality flags, shown first on the row so epidemiologists see them (see flagRecords). */
    flags: RecordFlag[];
};

export type FlagKind =
    | "TEST/INVALID"
    | "PARENT DELETED"
    | "PARENT MISSING"
    | "OUTSIDE EXTRACT"
    | "SURVEY MISMATCH"
    | "PARENT FLAGGED";

export type RecordFlag = { flag: FlagKind; detail: string };

/** What each flag means, for the _index legend. */
export const FLAG_MEANINGS: Record<FlagKind, string> = {
    "TEST/INVALID":
        "The record's Survey_id is missing or is not a survey in DHIS2: almost always test data.",
    "PARENT DELETED":
        "The record it belongs to (its case report or survey) was deleted in DHIS2; this record was not.",
    "PARENT MISSING": "The record it belongs to does not exist in DHIS2 (or no link was recorded).",
    "OUTSIDE EXTRACT":
        "The record it belongs to exists in DHIS2 but under another org unit, outside this extract.",
    "SURVEY MISMATCH": "Its own Survey_id differs from the Survey_id of the record it belongs to.",
    "PARENT FLAGGED": "The record it belongs to is itself flagged; see that record's flag.",
};

export type FormData = {
    form: ResolvedForm;
    records: Record_[];
    /** Ordered value column UIDs -> display label. */
    columns: Map<string, string>;
    /** Value column UID -> DHIS2 valueType; absent means the column is written as text. */
    valueTypes?: Map<string, string>;
};

const EXCEL_MAX_ROWS = 1_048_576;

/** Matches glass-dev's bulkDownloadAMUFiles.ts, which this server tolerates well. */
const FETCH_CONCURRENCY = 6;

/** Programs being extracted at once. Their page requests share REQUEST_CONCURRENCY below. */
const EXTRACT_CONCURRENCY = 3;

/**
 * Full-page data requests in flight across the whole run. A page carries every attribute,
 * event and value of up to --page-size records, far heavier than a discovery probe, so this
 * one shared cap is what keeps the load on DHIS2 bounded however pages and programs overlap.
 */
const REQUEST_CONCURRENCY = 4;

export type Limiter = <T>(task: () => Promise<T>) => Promise<T>;

/**
 * At most `limit` tasks run at once; the rest wait their turn in FIFO order. A finishing
 * task hands its slot straight to the next waiter, so the limit is never overshot.
 */
export function createLimiter(limit: number): Limiter {
    let active = 0;
    const waiting: (() => void)[] = [];

    return async task => {
        if (active < limit) active++;
        else await new Promise<void>(resolve => waiting.push(resolve));
        try {
            return await task();
        } finally {
            const next = waiting.shift();
            if (next) next();
            else active--;
        }
    };
}

/**
 * Runs `fn` over `items` with at most `limit` in flight, preserving input order in the
 * result. A pool, not fixed batches: a slow item holds up one slot, never the whole batch.
 * Bounded so a wide list can't open dozens of parallel requests against DHIS2 at once.
 * After a failure no further items are started.
 */
export async function mapWithConcurrency<T, R>(
    items: readonly T[],
    limit: number,
    fn: (item: T) => Promise<R>
): Promise<R[]> {
    const results = new Array<R>(items.length);
    let next = 0;

    const worker = async () => {
        while (next < items.length) {
            const index = next++;
            try {
                results[index] = await fn(items[index] as T);
            } catch (err) {
                next = items.length;
                throw err;
            }
        }
    };

    await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
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

/**
 * Fatal on failure: the datastore is what says which custom programs exist and which field
 * links a record to its parent, so carrying on without it would silently drop whole programs.
 */
async function fetchModules(api: D2Api, session: SessionManager): Promise<AMRSurveyModule[]> {
    try {
        const modules = await session.retryWithBackoff(() =>
            api.get<AMRSurveyModule[]>("/dataStore/amr-surveys/modules").getData()
        );
        if (!Array.isArray(modules)) throw new Error("unexpected content");
        return modules;
    } catch (err) {
        throw new Error(
            `Could not read datastore amr-surveys/modules (${
                err instanceof Error ? err.message : err
            }); ` + `custom forms and parent links cannot be resolved without it.`
        );
    }
}

/** The default program plus every custom variant any survey configures for it. */
export function candidateProgramIds(defaultUid: string, modules: AMRSurveyModule[]): string[] {
    const customs = modules
        .flatMap(module => Object.values(module.customForms ?? {}))
        .map(forms => forms[defaultUid])
        .filter((uid): uid is string => !!uid);
    return [...new Set([defaultUid, ...customs])];
}

async function resolveForms(
    api: D2Api,
    session: SessionManager,
    modules: AMRSurveyModule[],
    wanted: FormSpec[]
): Promise<{ resolved: ResolvedForm[]; unresolved: { spec: FormSpec; reason: string }[] }> {
    const { programs, dataSets } = await session.retryWithBackoff(() =>
        api.metadata
            .get({
                programs: { fields: { id: true, name: true, programType: true } },
                dataSets: { fields: { id: true, name: true } },
            })
            .getData()
    );

    const onServer = new Map<string, { name: string; kind: FormKind }>();
    for (const p of programs) {
        onServer.set(p.id, {
            name: p.name,
            kind: p.programType === "WITH_REGISTRATION" ? "tracker" : "event",
        });
    }
    for (const d of dataSets) onServer.set(d.id, { name: d.name, kind: "dataSet" });

    const resolved: ResolvedForm[] = [];
    const unresolved: { spec: FormSpec; reason: string }[] = [];

    for (const spec of wanted) {
        const matches = candidateProgramIds(spec.defaultUid, modules).flatMap(uid => {
            const found = onServer.get(uid);
            return found ? [{ uid, ...found }] : [];
        });

        if (matches.length === 0) {
            unresolved.push({
                spec,
                reason: `no program or dataSet with default UID ${spec.defaultUid} on this server`,
            });
            continue;
        }

        for (const match of matches) {
            resolved.push({
                ...spec,
                uid: match.uid,
                serverName: match.name,
                kind: spec.kind ?? match.kind,
                isCustom: match.uid !== spec.defaultUid,
                parentLinkField:
                    match.kind === "dataSet"
                        ? ""
                        : getParentDataElementForProgram(match.uid, modules),
                surveyLinkField: SURVEY_LINK_BY_DEFAULT_UID[spec.defaultUid] ?? "",
            });
        }
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

function reportResolution(
    resolved: ResolvedForm[],
    unresolved: { spec: FormSpec; reason: string }[]
) {
    console.log("\nResolved forms:");
    console.log(
        "  " +
            ["KEY".padEnd(18), "UID".padEnd(13), "KIND".padEnd(8), "CUSTOM".padEnd(7), "NAME"].join(
                ""
            )
    );
    for (const f of resolved) {
        console.log(
            "  " +
                [
                    f.key.padEnd(18),
                    f.uid.padEnd(13),
                    f.kind.padEnd(8),
                    (f.isCustom ? "yes" : "no").padEnd(7),
                    f.serverName,
                ].join("")
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
    stageById: Map<string, ProgramStageMeta>;
    /** dataElement/attribute id -> DHIS2 valueType; decides numeric and date cells. */
    valueTypes: Map<string, string>;
    /** dataElement/attribute id -> (option code -> option name), for option-set fields. */
    optionNames: Map<string, Map<string, string>>;
};

/** The option's name where the field is option-backed, else the stored value unchanged. */
function displayValue(meta: ProgramMeta, fieldId: string, value: string): string {
    return meta.optionNames.get(fieldId)?.get(value) ?? value;
}

/**
 * A field that holds values but is not on the current form: a data element since removed
 * from its stage (verified live: thousands of historical values on the case report form) or
 * an attribute no longer on the program. Its values are kept. Its column is labelled with
 * the field's full DHIS2 name, which, unlike the form text ("Route"), says which slot it was
 * ("Route1.1YES"), and is marked so nobody mistakes it for a current question.
 */
export type OffFormField = {
    kind: "dataElement" | "trackedEntityAttribute";
    id: string;
    /** Column sort position: just after the fields of the stage it was recorded on. */
    position: number;
    /** Label prefix ("Stage name: ") for main-sheet stage columns, else "". */
    prefix: string;
    /** Values seen, for the _index listing. */
    values: number;
    /** Known once its metadata is fetched (see labelOffFormFields). */
    valueType?: string;
};

/** Sheet ("" = the main sheet, else a repeatable stage id) -> column key -> off-form field. */
export type OffFormTracker = Map<string, Map<string, OffFormField>>;

const OFF_FORM_MARK = "[not on current form]";

/** Counts one value of an off-form column, registering the column on first sight. */
function noteOffForm(
    tracker: OffFormTracker | undefined,
    sheet: string,
    columnKey: string,
    field: Omit<OffFormField, "values">
): void {
    if (!tracker) return;
    let fields = tracker.get(sheet);
    if (!fields) tracker.set(sheet, (fields = new Map()));
    const known = fields.get(columnKey);
    if (known) known.values++;
    else fields.set(columnKey, { ...field, values: 1 });
}

/**
 * Program metadata is needed twice per form — once by discovery (to predict column counts
 * and find repeatable stages) and again by extraction (to label columns). These payloads
 * are large (one program here has ~500 data elements), so the result is memoised per run;
 * a run only ever fetches each program's metadata once.
 */
const programMetaCache = new Map<string, Promise<ProgramMeta>>();

function fetchProgramMeta(
    api: D2Api,
    session: SessionManager,
    programId: string
): Promise<ProgramMeta> {
    const cached = programMetaCache.get(programId);
    if (cached) return cached;
    // Evict on failure, so a failure that outlasts the retries doesn't poison every later
    // caller with the same rejected promise.
    const pending = session
        .retryWithBackoff(() => fetchProgramMetaUncached(api, programId))
        .catch(err => {
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

type ProgramField = {
    id: string;
    name: string;
    formName?: string;
    valueType?: string;
    optionSet?: { id: string };
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
                dataElements?: ProgramField[];
                trackedEntityAttributes?: ProgramField[];
                options?: { code: string; name: string; optionSet?: { id: string } }[];
            }>(`/programs/${programId}/metadata.json`, {
                fields: "programs,dataElements,trackedEntityAttributes,programTrackedEntityAttributes,options",
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
        placeSpeciesOtherAfterSpecies(dataElementLabels, dataElementOrder);

        const dataElementSection = new Map<string, string>();
        for (const section of stage.programStageSections ?? []) {
            for (const de of section.dataElements ?? [])
                dataElementSection.set(de.id, section.name);
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

    const namesByOptionSet = new Map<string, Map<string, string>>();
    for (const option of program.options ?? []) {
        const setId = option.optionSet?.id;
        if (!setId) continue;
        const names = namesByOptionSet.get(setId) ?? new Map<string, string>();
        names.set(option.code, option.name);
        namesByOptionSet.set(setId, names);
    }

    const valueTypes = new Map<string, string>();
    const optionNames = new Map<string, Map<string, string>>();
    for (const field of [
        ...(program.dataElements ?? []),
        ...(program.trackedEntityAttributes ?? []),
    ]) {
        if (field.valueType) valueTypes.set(field.id, field.valueType);
        const names = field.optionSet && namesByOptionSet.get(field.optionSet.id);
        if (names) optionNames.set(field.id, names);
    }

    return { attributes, attributeOrder, stageById, valueTypes, optionNames };
}

const SPECIES_LABEL = "Specify the species";
const SPECIES_OTHER_LABEL = "Species, other";

/**
 * On the Central Ref Lab form, the free-text "Species, other" (what the lab typed when the
 * species list had no match) is the last field of each species stage, ~120 columns after
 * the "Specify the species" it qualifies, with the antibiotic blocks in between. Anyone
 * reading the export sees "Other" and never reaches the text, so move it to sit directly
 * after its species. Mutates `order`; a stage lacking either field is left untouched.
 */
export function placeSpeciesOtherAfterSpecies(
    labels: Map<string, string>,
    order: Map<string, number>
): void {
    const idOf = (label: string) => [...labels].find(([, l]) => l === label)?.[0];
    const speciesId = idOf(SPECIES_LABEL);
    const otherId = idOf(SPECIES_OTHER_LABEL);
    const speciesPosition = speciesId === undefined ? undefined : order.get(speciesId);
    if (otherId === undefined || speciesPosition === undefined) return;

    order.set(otherId, speciesPosition + 0.5);
}

/**
 * The label a column gets, before de-duplication: the form's own section name where the
 * stage defines one (e.g. "S1 - Antibiotic 1"), otherwise the stage name.
 */
function stageColumnLabel(
    stage: ProgramStageMeta,
    dataElementId: string,
    fieldLabel: string
): string {
    return `${stage.dataElementSection.get(dataElementId) ?? stage.name}: ${fieldLabel}`;
}

/**
 * Marks a main-sheet column as coming from a program-stage data element rather than a
 * tracked-entity attribute, so the two id spaces can share one column map.
 */
const STAGE_COLUMN_PREFIX = "stage:";

/** Where a column belongs in form order. Attributes first, then stage data in stage order. */
const STAGE_COLUMN_OFFSET = 1_000_000;

/** Sorts after every field of `stage` (or, with no stage, after every attribute). */
function afterStagePosition(stage: ProgramStageMeta | undefined): number {
    return STAGE_COLUMN_OFFSET * (1 + (stage ? 1 + stage.sortOrder : 0)) - 1;
}

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
 * Puts columns into form order. Order matters for reading: the Central Ref Lab form repeats
 * a 4-field block per antibiotic, and without this the antibiotic and its own AST result do
 * not end up side by side. Off-form columns go after their stage's fields, by label.
 */
export function orderColumns(
    columns: Map<string, string>,
    meta: ProgramMeta,
    offForm?: Map<string, OffFormField>
): Map<string, string> {
    const position = (key: string) => offForm?.get(key)?.position ?? columnSortKey(key, meta);
    return new Map(
        [...columns.entries()].sort(
            ([a, labelA], [b, labelB]) =>
                position(a) - position(b) ||
                labelA.localeCompare(labelB, undefined, { numeric: true })
        )
    );
}

/**
 * Makes every header unique, keeping order. A repeated block reuses the same field names — a
 * bare "Specify the antibiotic" appearing 24 times tells the reader nothing — and two
 * programs merged into one sheet can each bring a field with the same label. Where a label
 * still repeats, occurrences are numbered in column order. Applied once, at write time, so it
 * sees the final (merged) column set.
 */
export function uniqueLabels(columns: Map<string, string>): Map<string, string> {
    const labelCounts = new Map<string, number>();
    for (const label of columns.values()) labelCounts.set(label, (labelCounts.get(label) ?? 0) + 1);

    const seen = new Map<string, number>();
    const result = new Map<string, string>();
    for (const [key, label] of columns) {
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
    /** Discovered record count of the program being fetched, when known: enables parallel pages. */
    total?: number;
    /** The run's shared cap on full-page data requests. */
    requests: Limiter;
    /** Org units to read dataSets under when no --org-unit/--country scope is given. */
    dataSetRoots?: string[];
};

/**
 * Meta-column-key lists, named once so the discovery/forecast column-count math (which
 * predicts a sheet's shape before any records are fetched) can never drift from what the
 * real fetch functions below actually write — both reference the same `.length`.
 */
const TRACKER_MAIN_META_KEYS = [
    "created_at",
    "updated_at",
    "enrolled_at",
    "status",
    "program",
] as const;
const TRACKER_STAGE_CHILD_META_KEYS = ["occurred_at", "status", "program"] as const;
const EVENT_META_KEYS = ["occurred_at", "created_at", "updated_at", "status", "program"] as const;
/** flag, flag_detail, path, record_id, org_unit_id, org_unit_name — on every sheet regardless of kind. */
const FIXED_NON_META_COLUMNS = 6;

/**
 * The ResolvedForm shape a repeatable stage's child sheet gets. Shared by fetchTracker
 * (which creates these for real) and discovery's forecast (which previews them before any
 * extraction runs), so a predicted sheet can never describe a shape extraction wouldn't
 * actually produce.
 */
export function buildStageChildForm(
    owner: ResolvedForm,
    stage: ProgramStageMeta,
    childKey: string
): ResolvedForm {
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
export function assignStageChildKey(
    ownerKey: string,
    stage: ProgramStageMeta,
    usedKeys: Set<string>
): string {
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
 * Guards the paging loops. Returns the items of a page not seen on an earlier page, so a
 * record that shifts between pages while the data changes is never written twice. Returns
 * undefined once a page adds nothing new: a server that ignores `page` would otherwise
 * return the same page forever.
 */
export function makePageGuard(formKey: string) {
    const seen = new Set<string>();

    return function fresh<T>(items: T[], idOf: (item: T) => string): T[] | undefined {
        const added = items.filter(item => !seen.has(idOf(item)));
        added.forEach(item => seen.add(idOf(item)));

        if (items.length > 0 && added.length === 0) {
            console.warn(
                `\n  ! ${formKey}: a page returned no new records; stopping to avoid a paging loop.`
            );
            return undefined;
        }
        return added;
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
 * pattern the rest of this file uses. Values of fields no longer on the form are kept and
 * noted in `offForm` (see OffFormField).
 */
export function buildTrackerRecords(
    tei: TrackedEntityInput,
    form: ResolvedForm,
    meta: ProgramMeta,
    columns: Map<string, string>,
    stageColumns: Map<string, Map<string, string>>,
    report?: IntegrityReport,
    offForm?: OffFormTracker
): { main: Record_; stageRecords: StageChildRecord[] } {
    const values = new Map<string, string>();
    for (const attr of tei.attributes ?? []) {
        if (attr.value === undefined || attr.value === null) continue;
        values.set(attr.attribute, displayValue(meta, attr.attribute, String(attr.value)));
        const label = meta.attributes.get(attr.attribute);
        if (label === undefined) {
            noteOffForm(offForm, "", attr.attribute, {
                kind: "trackedEntityAttribute",
                id: attr.attribute,
                position: afterStagePosition(undefined),
                prefix: "",
            });
        }
        if (!columns.has(attr.attribute)) columns.set(attr.attribute, label ?? attr.attribute);
    }

    const stageRecords: StageChildRecord[] = [];
    const filledStages = new Set<string>();

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
            // A single-entry stage's values share one row, so a second event with data
            // overwrites the first's. Not silent: counted and surfaced in _index.
            if ((event.dataValues?.length ?? 0) > 0) {
                if (filledStages.has(stage.id) && report) {
                    incrementStageCount(report.duplicateStageEvents, form.key, stage.id);
                }
                filledStages.add(stage.id);
            }
            for (const dv of event.dataValues ?? []) {
                if (dv.value === undefined || dv.value === null) continue;
                const key = `${STAGE_COLUMN_PREFIX}${dv.dataElement}`;
                values.set(key, displayValue(meta, dv.dataElement, String(dv.value)));
                const label = stage.dataElementLabels.get(dv.dataElement);
                if (label === undefined) {
                    noteOffForm(offForm, "", key, {
                        kind: "dataElement",
                        id: dv.dataElement,
                        position: afterStagePosition(stage),
                        prefix: `${stage.name}: `,
                    });
                }
                if (!columns.has(key)) {
                    columns.set(
                        key,
                        stageColumnLabel(stage, dv.dataElement, label ?? dv.dataElement)
                    );
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
            ["program", form.serverName],
        ]),
        flags: [],
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
            childValues.set(dv.dataElement, displayValue(meta, dv.dataElement, String(dv.value)));
            const label = stage.dataElementLabels.get(dv.dataElement);
            if (label === undefined) {
                noteOffForm(offForm, stage.id, dv.dataElement, {
                    kind: "dataElement",
                    id: dv.dataElement,
                    position: afterStagePosition(stage),
                    prefix: "",
                });
            }
            if (!childColumns.has(dv.dataElement))
                childColumns.set(dv.dataElement, label ?? dv.dataElement);
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
                    ["program", form.serverName],
                ]),
                flags: [],
            },
        });
    }

    return { main, stageRecords };
}

/** Ids per metadata request: keeps `id:in:[...]` URLs well under server limits. */
const ID_BATCH = 100;

type FieldInfo = { name: string; valueType?: string; options: Map<string, string> };

/** Full name, valueType and option names of the given fields. */
async function fetchFieldInfo(
    api: D2Api,
    session: SessionManager,
    kind: OffFormField["kind"],
    ids: string[]
): Promise<Map<string, FieldInfo>> {
    const endpoint = kind === "dataElement" ? "dataElements" : "trackedEntityAttributes";
    const info = new Map<string, FieldInfo>();

    for (let i = 0; i < ids.length; i += ID_BATCH) {
        const batch = ids.slice(i, i + ID_BATCH);
        const resp = await session.retryWithBackoff(() =>
            api
                .get<{
                    [endpoint: string]:
                        | {
                              id: string;
                              name: string;
                              valueType?: string;
                              optionSet?: { options?: { code: string; name: string }[] };
                          }[]
                        | undefined;
                }>(`/${endpoint}.json`, {
                    filter: `id:in:[${batch.join(",")}]`,
                    fields: "id,name,valueType,optionSet[options[code,name]]",
                    paging: false,
                })
                .getData()
        );
        for (const field of resp[endpoint] ?? []) {
            info.set(field.id, {
                name: field.name,
                valueType: field.valueType,
                options: new Map((field.optionSet?.options ?? []).map(o => [o.code, o.name])),
            });
        }
    }
    return info;
}

type SheetInProgress = { columns: Map<string, string>; records: Record_[] };

/**
 * Labels the off-form columns of one program's sheets with each field's full DHIS2 name,
 * turns their option codes into names (the program's own metadata no longer covers them)
 * and records their valueType. A field whose metadata is gone entirely keeps its id as the
 * label and is reported in report.unresolvedHeaders.
 */
async function labelOffFormFields(
    api: D2Api,
    session: SessionManager,
    formKey: string,
    tracker: OffFormTracker,
    sheets: Map<string, SheetInProgress>,
    report: IntegrityReport
): Promise<void> {
    const fields = [...tracker.values()].flatMap(byColumn => [...byColumn.values()]);
    if (fields.length === 0) return;

    const idsOf = (kind: OffFormField["kind"]) => [
        ...new Set(fields.filter(f => f.kind === kind).map(f => f.id)),
    ];
    const [dataElements, attributes] = await Promise.all([
        fetchFieldInfo(api, session, "dataElement", idsOf("dataElement")),
        fetchFieldInfo(api, session, "trackedEntityAttribute", idsOf("trackedEntityAttribute")),
    ]);

    for (const [sheet, byColumn] of tracker) {
        const target = sheets.get(sheet);
        if (!target) continue;

        for (const [columnKey, field] of byColumn) {
            const info = (field.kind === "dataElement" ? dataElements : attributes).get(field.id);
            if (!info) {
                target.columns.set(columnKey, `${field.prefix}${field.id} [deleted field]`);
                report.unresolvedHeaders.push({ formKey, id: field.id, kind: field.kind });
                continue;
            }
            target.columns.set(columnKey, `${field.prefix}${info.name} ${OFF_FORM_MARK}`);
            field.valueType = info.valueType;
            if (info.options.size === 0) continue;
            for (const record of target.records) {
                const value = record.values.get(columnKey);
                const name = value === undefined ? undefined : info.options.get(value);
                if (name !== undefined) record.values.set(columnKey, name);
            }
        }
    }
}

/** Lists a sheet's off-form columns, with their value counts, for _index. */
function reportOffForm(
    report: IntegrityReport,
    sheet: string,
    fields: Map<string, OffFormField> | undefined,
    columns: Map<string, string>
): void {
    for (const [columnKey, field] of fields ?? []) {
        report.offFormFields.push({
            sheet,
            column: columns.get(columnKey) ?? columnKey,
            values: field.values,
        });
    }
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

/** The valueType of each column, so the writer can emit real numbers and dates. */
function columnValueTypes(
    columns: Map<string, string>,
    meta: ProgramMeta,
    offForm?: Map<string, OffFormField>
): Map<string, string> {
    const types = new Map<string, string>();
    for (const key of columns.keys()) {
        const fieldId = key.startsWith(STAGE_COLUMN_PREFIX)
            ? key.slice(STAGE_COLUMN_PREFIX.length)
            : key;
        const type = offForm?.get(key)?.valueType ?? meta.valueTypes.get(fieldId);
        if (type) types.set(key, type);
    }
    return types;
}

/**
 * Every record of a paged endpoint, in page order, each once. With a known total (from
 * discovery) every page is requested at once and the shared request limiter paces them;
 * without one, and past the known pages if records were added since discovery, pages are
 * read one at a time until a short page. A record that shifts between pages while the data
 * changes is kept once; reconciliation reports any change in count.
 */
export async function fetchAllPages<T>(
    formKey: string,
    pageSize: number,
    total: number | undefined,
    fetchPage: (page: number) => Promise<T[]>,
    idOf: (item: T) => string
): Promise<T[]> {
    const fresh = makePageGuard(formKey);
    const items: T[] = [];
    const take = (page: T[]): boolean => {
        const added = fresh(page, idOf);
        if (added) items.push(...added);
        return added !== undefined;
    };

    const known = total ? Math.ceil(total / pageSize) : 0;
    const pages = await Promise.all(Array.from({ length: known }, (_, i) => fetchPage(i + 1)));
    pages.forEach(take);

    let last = pages[pages.length - 1];
    for (let page = known + 1; !last || last.length === pageSize; page++) {
        last = await fetchPage(page);
        if (!take(last)) break;
    }
    return items;
}

async function fetchTracker(
    api: D2Api,
    form: ResolvedForm,
    meta: ProgramMeta,
    opts: FetchOpts,
    session: SessionManager,
    report: IntegrityReport
): Promise<{ main: FormData; stageForms: FormData[] }> {
    const teis = await fetchAllPages(
        form.key,
        opts.pageSize,
        opts.total,
        page =>
            opts
                .requests(() =>
                    session.retryWithBackoff(() =>
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
                                // An explicit order keeps pages stable while records are added.
                                order: [{ type: "field", field: "createdAt", direction: "asc" }],
                                page,
                                pageSize: opts.pageSize,
                                totalPages: false,
                            })
                            .getData()
                    )
                )
                .then(resp => resp.instances ?? []),
        tei => tei.trackedEntity
    );

    const records: Record_[] = [];
    const columns = new Map<string, string>();
    // stage id -> child rows / columns of a repeatable stage.
    const stageRows = new Map<string, Record_[]>();
    const stageColumns = new Map<string, Map<string, string>>();
    const offForm: OffFormTracker = new Map();

    for (const tei of teis) {
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
            offForm
        );
        records.push(main);
        for (const { stageId, record } of stageRecords) {
            const rows = stageRows.get(stageId);
            if (rows) rows.push(record);
            else stageRows.set(stageId, [record]);
        }
    }
    console.log(`  ${form.key}: ${records.length} records`);

    const sheets = new Map<string, SheetInProgress>([["", { columns, records }]]);
    for (const [stageId, childColumns] of stageColumns) {
        sheets.set(stageId, { columns: childColumns, records: stageRows.get(stageId) ?? [] });
    }
    await labelOffFormFields(api, session, form.key, offForm, sheets, report);

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
        const childOffForm = offForm.get(stageId);
        const unordered = stageColumns.get(stageId) ?? new Map<string, string>();
        reportOffForm(report, childKey, childOffForm, unordered);
        const childColumns = orderColumns(unordered, meta, childOffForm);
        stageForms.push({
            form: buildStageChildForm(form, stage, childKey),
            records: rows,
            columns: childColumns,
            valueTypes: columnValueTypes(childColumns, meta, childOffForm),
        });
    }

    reportOffForm(report, form.key, offForm.get(""), columns);
    const mainColumns = orderColumns(columns, meta, offForm.get(""));
    return {
        main: {
            form,
            records,
            columns: mainColumns,
            valueTypes: columnValueTypes(mainColumns, meta, offForm.get("")),
        },
        stageForms,
    };
}

async function fetchEvents(
    api: D2Api,
    form: ResolvedForm,
    meta: ProgramMeta,
    opts: FetchOpts,
    session: SessionManager,
    report: IntegrityReport
): Promise<FormData> {
    const events = await fetchAllPages(
        form.key,
        opts.pageSize,
        opts.total,
        page =>
            opts
                .requests(() =>
                    session.retryWithBackoff(() =>
                        api.tracker.events
                            .get({
                                fields: {
                                    event: true,
                                    programStage: true,
                                    orgUnit: true,
                                    occurredAt: true,
                                    createdAt: true,
                                    updatedAt: true,
                                    status: true,
                                    dataValues: { dataElement: true, value: true },
                                },
                                program: form.uid,
                                ...ouParams(opts.orgUnit),
                                order: "createdAt:asc",
                                page,
                                pageSize: opts.pageSize,
                                totalPages: false,
                            })
                            .getData()
                    )
                )
                .then(resp => resp.instances ?? []),
        event => event.event
    );

    const columns = new Map<string, string>();
    const offForm: OffFormTracker = new Map();
    const records: Record_[] = events.map(event => {
        const stage = meta.stageById.get(event.programStage);
        const values = new Map<string, string>();
        for (const dv of event.dataValues ?? []) {
            if (dv.value === undefined || dv.value === null) continue;
            values.set(dv.dataElement, displayValue(meta, dv.dataElement, String(dv.value)));
            const label = stage?.dataElementLabels.get(dv.dataElement);
            if (label === undefined) {
                noteOffForm(offForm, "", dv.dataElement, {
                    kind: "dataElement",
                    id: dv.dataElement,
                    position: afterStagePosition(stage),
                    prefix: "",
                });
            }
            if (!columns.has(dv.dataElement)) columns.set(dv.dataElement, label ?? dv.dataElement);
        }

        return {
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
                ["program", form.serverName],
            ]),
            flags: [],
        };
    });
    console.log(`  ${form.key}: ${records.length} records`);

    await labelOffFormFields(
        api,
        session,
        form.key,
        offForm,
        new Map([["", { columns, records }]]),
        report
    );
    reportOffForm(report, form.key, offForm.get(""), columns);
    const ordered = orderColumns(columns, meta, offForm.get(""));
    return {
        form,
        records,
        columns: ordered,
        valueTypes: columnValueTypes(ordered, meta, offForm.get("")),
    };
}

/** dataValueSets has no cheap probe, so without explicit dates the whole history is read. */
const DATASET_DEFAULT_START_DATE = "2000-01-01";

const NUMERIC_VALUE_TYPES = new Set([
    "NUMBER",
    "INTEGER",
    "INTEGER_POSITIVE",
    "INTEGER_NEGATIVE",
    "INTEGER_ZERO_OR_POSITIVE",
    "PERCENTAGE",
    "UNIT_INTERVAL",
]);

type NamedRow = { id: string; name: string };

/** Names for everything a dataValue of this dataSet refers to by id. */
async function fetchDataSetNames(
    api: D2Api,
    session: SessionManager,
    uid: string
): Promise<{
    dataElements: Map<string, { label: string; valueType?: string }>;
    optionCombos: Map<string, string>;
    /** The dataSet's own attribute option combos (its wards): which values are its own. */
    attributeOptionCombos: Set<string>;
}> {
    const ds = await session.retryWithBackoff(() =>
        api
            .get<{
                categoryCombo?: { categoryOptionCombos?: NamedRow[] };
                dataSetElements?: {
                    dataElement: NamedRow & {
                        formName?: string;
                        valueType?: string;
                        categoryCombo?: { categoryOptionCombos?: NamedRow[] };
                    };
                }[];
            }>(`/dataSets/${uid}.json`, {
                fields:
                    "categoryCombo[categoryOptionCombos[id,name]]," +
                    "dataSetElements[dataElement[id,name,formName,valueType,categoryCombo[categoryOptionCombos[id,name]]]]",
            })
            .getData()
    );

    const dataElements = new Map<string, { label: string; valueType?: string }>();
    const optionCombos = new Map<string, string>();
    for (const coc of ds.categoryCombo?.categoryOptionCombos ?? [])
        optionCombos.set(coc.id, coc.name);
    for (const { dataElement: de } of ds.dataSetElements ?? []) {
        dataElements.set(de.id, { label: de.formName || de.name, valueType: de.valueType });
        for (const coc of de.categoryCombo?.categoryOptionCombos ?? [])
            optionCombos.set(coc.id, coc.name);
    }
    const attributeOptionCombos = new Set(
        (ds.categoryCombo?.categoryOptionCombos ?? []).map(coc => coc.id)
    );
    return { dataElements, optionCombos, attributeOptionCombos };
}

/**
 * The ward statistics are aggregate dataSets, not programs: one row per dataValue, keyed by
 * data element, period, org unit (the facility) and the two option combos. The attribute
 * option combo identifies the ward (and, for Ward Summary Statistics, the specialty).
 *
 * The two ward dataSets share their data elements, and dataValueSets selects by data
 * element, so each request returns both dataSets' values (verified live: the same 230 values
 * twice). A value belongs to the dataSet whose attribute combo holds its ward; the others
 * are left to that dataSet's own sheet and counted in report.otherDataSetValues.
 */
async function fetchDataSet(
    api: D2Api,
    form: ResolvedForm,
    opts: FetchOpts,
    session: SessionManager,
    report: IntegrityReport
): Promise<FormData> {
    const orgUnits = opts.orgUnit ? [opts.orgUnit] : opts.dataSetRoots ?? [];
    const columns = new Map<string, string>([
        ["period", "Period"],
        ["ward", "Ward"],
        ["data_element", "Data element"],
        ["disaggregation", "Disaggregation"],
        ["value", "Value"],
    ]);
    if (orgUnits.length === 0) {
        console.warn(`  ! ${form.key}: skipped (no org unit to read it under).`);
        return { form, records: [], columns };
    }

    const [resp, names] = await Promise.all([
        opts.requests(() =>
            session.retryWithBackoff(() =>
                api.dataValues
                    .getSet({
                        dataSet: [form.uid],
                        orgUnit: orgUnits,
                        startDate: opts.startDate ?? DATASET_DEFAULT_START_DATE,
                        endDate: opts.endDate ?? new Date().toISOString().slice(0, 10),
                        children: true,
                    })
                    .getData()
            )
        ),
        fetchDataSetNames(api, session, form.uid),
    ]);

    const nameOf = (id: string | undefined) => (id ? names.optionCombos.get(id) ?? id : "");
    const dataValues = resp.dataValues ?? [];
    const own = dataValues.filter(dv =>
        names.attributeOptionCombos.has(dv.attributeOptionCombo ?? "")
    );
    if (own.length < dataValues.length) {
        report.otherDataSetValues.push({
            formKey: form.key,
            count: dataValues.length - own.length,
        });
    }
    const records: Record_[] = own.map(dv => {
        const coc = nameOf(dv.categoryOptionCombo);
        return {
            id: `${dv.dataElement}-${dv.period}-${dv.orgUnit}-${dv.categoryOptionCombo}-${dv.attributeOptionCombo}`,
            parentId: "",
            surveyId: "",
            facilityId: "",
            orgUnit: dv.orgUnit ?? "",
            label: "",
            values: new Map<string, string>([
                ["period", String(dv.period ?? "")],
                ["ward", nameOf(dv.attributeOptionCombo)],
                ["data_element", names.dataElements.get(dv.dataElement)?.label ?? dv.dataElement],
                ["disaggregation", coc === "default" ? "" : coc],
                ["value", String(dv.value ?? "")],
            ]),
            meta: new Map<string, string>([
                ["last_updated", dv.lastUpdated ?? ""],
                ["stored_by", dv.storedBy ?? ""],
            ]),
            flags: [],
        };
    });

    const allNumeric = [...names.dataElements.values()].every(
        de => de.valueType !== undefined && NUMERIC_VALUE_TYPES.has(de.valueType)
    );
    const others = dataValues.length - own.length;
    console.log(
        `  ${form.key}: ${records.length} data values` +
            (others > 0 ? ` (${others} more belong to the other ward dataSet's wards)` : "")
    );
    return { form, records, columns, valueTypes: new Map(allNumeric ? [["value", "NUMBER"]] : []) };
}

// --- Merging programs ---------------------------------------------------------

const joinUnique = (a: string, b: string) =>
    a === b ? a : [...new Set([...a.split("; "), ...b.split("; ")])].join("; ");

/**
 * A form can be backed by several programs: the default plus each custom variant. They share
 * most data elements, so their rows are stacked into ONE sheet per form key; the `program`
 * column says which program a row came from. Columns are the union in first-seen order, so
 * the default program's order leads. Repeatable-stage child sheets merge the same way.
 */
export function mergeByFormKey(all: FormData[]): FormData[] {
    const merged = new Map<string, FormData>();

    for (const data of all) {
        const prior = merged.get(data.form.key);
        if (!prior) {
            merged.set(data.form.key, data);
            continue;
        }

        const columns = new Map(prior.columns);
        for (const [key, label] of data.columns) if (!columns.has(key)) columns.set(key, label);

        merged.set(data.form.key, {
            form: {
                ...prior.form,
                uid: joinUnique(prior.form.uid, data.form.uid),
                serverName: joinUnique(prior.form.serverName, data.form.serverName),
                parentLinkField: joinUnique(prior.form.parentLinkField, data.form.parentLinkField),
                isCustom: prior.form.isCustom || data.form.isCustom,
            },
            records: prior.records.concat(data.records),
            columns,
            valueTypes: new Map([...(prior.valueTypes ?? []), ...(data.valueTypes ?? [])]),
        });
    }

    return [...merged.values()];
}

/**
 * Expected record count per sheet key: the discovered totals of the programs actually
 * extracted for it. A key is absent when any of its programs has no known total (dataSets,
 * runs without a scope), so reconciliation never compares against a partial number.
 */
export function expectedRecordCounts(
    extracted: ResolvedForm[],
    totalByUid: Map<string, number | undefined>,
    discoveries: FormDiscovery[] = []
): Map<string, number> {
    const sums = new Map<string, number>();
    const unknown = new Set<string>();

    for (const form of extracted) {
        const total = totalByUid.get(form.uid);
        if (total === undefined) unknown.add(form.key);
        else sums.set(form.key, (sums.get(form.key) ?? 0) + total);
    }
    for (const key of unknown) sums.delete(key);

    // Repeatable-stage sheets: discovery probed their row counts too.
    const extractedUids = new Set(extracted.map(f => f.uid));
    for (const d of discoveries) {
        if (!extractedUids.has(d.form.uid)) continue;
        for (const s of d.stages) sums.set(s.sheetKey, (sums.get(s.sheetKey) ?? 0) + s.rows);
    }
    return sums;
}

// --- Org unit names ------------------------------------------------------------

async function fetchOrgUnitNames(
    api: D2Api,
    session: SessionManager,
    ids: string[]
): Promise<Map<string, string>> {
    const names = new Map<string, string>();
    const chunkSize = 200;

    for (let i = 0; i < ids.length; i += chunkSize) {
        const chunk = ids.slice(i, i + chunkSize);
        const { organisationUnits } = await session.retryWithBackoff(() =>
            api.metadata
                .get({
                    organisationUnits: {
                        fields: { id: true, name: true },
                        filter: { id: { in: chunk } },
                    },
                })
                .getData()
        );
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
                const id =
                    attempt === 0
                        ? shortStableId(source.stableId)
                        : `${shortStableId(source.stableId)}${attempt}`;
                const suffix = `~${id}`;
                candidate =
                    safeLabel.slice(0, Math.max(0, EXCEL_SHEET_NAME_MAX - suffix.length)) + suffix;
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

/**
 * Run-level findings for _index and the console. Record-level problems (orphans, test data,
 * survey mismatches) are not here: they are flags on the records themselves (flagRecords).
 */
export type IntegrityReport = {
    /** (survey, orgUnit) pairs matching more than one Facility record. */
    ambiguousFacilities: { surveyId: string; orgUnit: string; facilityIds: string[] }[];
    /** Columns of fields no longer on the form, with how many values they hold. */
    offFormFields: { sheet: string; column: string; values: number }[];
    /** dataValues a dataSet request returned that belong to another dataSet's wards (see fetchDataSet). */
    otherDataSetValues: { formKey: string; count: number }[];
    /** Records for which no Facility could be derived. */
    unresolvedFacility: { formKey: string; count: number }[];
    /** Stage events whose programStage id did not resolve in program metadata (data was skipped). */
    unresolvedStageEvents: { formKey: string; programStage: string; count: number }[];
    /** Extra events (with data) on a single-entry stage: their values overwrote the earlier event's in the shared row. */
    duplicateStageEvents: { formKey: string; programStage: string; count: number }[];
    /** Sheet names that needed a disambiguation suffix to stay unique. */
    sheetNameCollisions: { key: string; assignedName: string }[];
    /** Off-form fields whose metadata no longer exists at all: the header shows the raw id. */
    unresolvedHeaders: {
        formKey: string;
        id: string;
        kind: "dataElement" | "trackedEntityAttribute";
    }[];
};

export function emptyIntegrityReport(): IntegrityReport {
    return {
        ambiguousFacilities: [],
        offFormFields: [],
        otherDataSetValues: [],
        unresolvedFacility: [],
        unresolvedStageEvents: [],
        duplicateStageEvents: [],
        sheetNameCollisions: [],
        unresolvedHeaders: [],
    };
}

function incrementStageCount(
    counts: IntegrityReport["unresolvedStageEvents"],
    formKey: string,
    programStage: string
): void {
    const existing = counts.find(u => u.formKey === formKey && u.programStage === programStage);
    if (existing) existing.count += 1;
    else counts.push({ formKey, programStage, count: 1 });
}

/** Records one occurrence of an event whose programStage id did not resolve in metadata. */
export function recordUnresolvedStageEvent(
    report: IntegrityReport,
    formKey: string,
    programStage: string
): void {
    incrementStageCount(report.unresolvedStageEvents, formKey, programStage);
}

export function reportIntegrity(report: IntegrityReport, all: FormData[]): void {
    const {
        ambiguousFacilities,
        offFormFields,
        unresolvedFacility,
        unresolvedStageEvents,
        duplicateStageEvents,
        sheetNameCollisions,
        unresolvedHeaders,
    } = report;
    const flagged = summariseFlags(all);
    if (
        flagged.length === 0 &&
        ambiguousFacilities.length === 0 &&
        offFormFields.length === 0 &&
        unresolvedFacility.length === 0 &&
        unresolvedStageEvents.length === 0 &&
        duplicateStageEvents.length === 0 &&
        sheetNameCollisions.length === 0 &&
        unresolvedHeaders.length === 0
    ) {
        console.log(
            "  integrity: no flagged records, no off-form fields, no skipped or overwritten events."
        );
        return;
    }

    console.warn("\n  ! Integrity findings (see the _index sheet for detail):");
    for (const f of flagged)
        console.warn(`      ${f.sheet}: ${f.count} record(s) flagged ${f.flag}`);
    if (offFormFields.length > 0) {
        const values = offFormFields.reduce((sum, f) => sum + f.values, 0);
        console.warn(
            `      ${offFormFields.length} column(s) of fields no longer on the form, ${values} value(s) kept`
        );
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
    if (unresolvedStageEvents.length > 0) {
        for (const u of unresolvedStageEvents) {
            console.warn(
                `      ${u.formKey}: ${u.count} event(s) on unresolvable stage ${u.programStage} SKIPPED (data lost)`
            );
        }
    }
    for (const u of duplicateStageEvents) {
        console.warn(
            `      ${u.formKey}: ${u.count} extra event(s) on single-entry stage ${u.programStage}; ` +
                `their values overwrote the earlier event's`
        );
    }
    if (sheetNameCollisions.length > 0) {
        console.warn(`      sheet name collisions resolved: ${sheetNameCollisions.length}`);
    }
    if (unresolvedHeaders.length > 0) {
        console.warn(
            `      fields deleted from DHIS2 metadata (header shows the raw id): ${unresolvedHeaders.length}`
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
        if (unresolved > 0)
            report.unresolvedFacility.push({ formKey: data.form.key, count: unresolved });
    }
}

/** What DHIS2 says about an id that records point at but this extract does not contain. */
export type ReferenceStatus =
    | { status: "deleted"; lastUpdated: string }
    /** Exists and is not deleted, under `orgUnit` (a display name). */
    | { status: "elsewhere"; orgUnit: string };

/** Looks up ids in DHIS2; an id missing from the result does not exist there. */
export type ReferenceLookup = (ids: {
    surveys: string[];
    trackedEntities: string[];
}) => Promise<Map<string, ReferenceStatus>>;

const isUid = (id: string) => UID_RE.test(id);

function referenceFlag(
    what: string,
    id: string,
    status: ReferenceStatus | undefined,
    isSurvey: boolean
): RecordFlag {
    if (status?.status === "deleted") {
        return {
            flag: "PARENT DELETED",
            detail: `${what} ${id} was deleted in DHIS2 (last changed ${status.lastUpdated.slice(
                0,
                10
            )})`,
        };
    }
    if (status?.status === "elsewhere") {
        return {
            flag: "OUTSIDE EXTRACT",
            detail: `${what} ${id} exists in DHIS2 under ${status.orgUnit}, outside this extract`,
        };
    }
    return isSurvey
        ? { flag: "TEST/INVALID", detail: `Survey_id "${id}" is not a survey in DHIS2` }
        : { flag: "PARENT MISSING", detail: `${what} "${id}" does not exist in DHIS2` };
}

function addFlag(record: Record_, flag: RecordFlag): void {
    if (!record.flags.some(f => f.flag === flag.flag && f.detail === flag.detail))
        record.flags.push(flag);
}

/**
 * Flags the records epidemiologists should look at twice: test or invalid survey ids,
 * records whose parent was deleted in DHIS2 (children are not deleted with it), parents that
 * never existed or sit outside the extract, survey mismatches, and records whose parent is
 * itself flagged. Every dangling id is looked up in one batch, so "deleted" is told apart
 * from "never existed". Parents are flagged before their children, so flags propagate down.
 *
 * The Survey_id check needs the Survey form in the run (without it every id would look
 * unknown); a parent link is checked only when the parent form was extracted.
 */
export async function flagRecords(
    all: FormData[],
    recordIndex: Map<string, Map<string, Record_>>,
    lookUp: ReferenceLookup
): Promise<void> {
    const byKey = new Map(all.map(d => [d.form.key, d]));
    const survey = all.find(d => d.form.defaultUid === PREVALENCE_SURVEY_FORM_ID);
    const surveys = survey && recordIndex.get(survey.form.key);
    const depth = (data: FormData): number => {
        const parent = data.form.parentKey ? byKey.get(data.form.parentKey) : undefined;
        return parent && parent !== data ? 1 + depth(parent) : 0;
    };

    const checks = all
        .filter(data => data.form.kind !== "dataSet" && data !== survey)
        .map(data => {
            const parent = data.form.parentKey ? byKey.get(data.form.parentKey) : undefined;
            return {
                data,
                // A repeatable-stage row carries its owner's Survey_id; the owner's check covers it.
                checkSurvey: !!surveys && data.form.kind !== "trackerStage",
                // Where the parent IS the Survey, the Survey_id check already covers the link.
                parent: parent && parent !== survey ? parent : undefined,
            };
        })
        .sort((a, b) => depth(a.data) - depth(b.data));

    const danglingSurveys = new Set<string>();
    const danglingParents = new Set<string>();
    for (const { data, checkSurvey, parent } of checks) {
        const parents = parent && recordIndex.get(parent.form.key);
        for (const record of data.records) {
            if (checkSurvey && record.surveyId && !surveys?.has(record.surveyId)) {
                danglingSurveys.add(record.surveyId);
            }
            if (parents && record.parentId && !parents.has(record.parentId))
                danglingParents.add(record.parentId);
        }
    }
    const statuses =
        danglingSurveys.size + danglingParents.size > 0
            ? await lookUp({
                  surveys: [...danglingSurveys].filter(isUid),
                  trackedEntities: [...danglingParents].filter(isUid),
              })
            : new Map<string, ReferenceStatus>();

    for (const { data, checkSurvey, parent } of checks) {
        const parents = parent && recordIndex.get(parent.form.key);
        for (const record of data.records) {
            if (checkSurvey) {
                if (!record.surveyId)
                    addFlag(record, { flag: "TEST/INVALID", detail: "No Survey_id recorded" });
                else if (!surveys?.has(record.surveyId)) {
                    addFlag(
                        record,
                        referenceFlag(
                            "Survey",
                            record.surveyId,
                            statuses.get(record.surveyId),
                            true
                        )
                    );
                }
            }
            if (!parent || !parents) continue;

            const parentKey = parent.form.key;
            const parentRecord = record.parentId ? parents.get(record.parentId) : undefined;
            if (!record.parentId) {
                addFlag(record, { flag: "PARENT MISSING", detail: `No ${parentKey} id recorded` });
            } else if (!parentRecord) {
                addFlag(
                    record,
                    referenceFlag(parentKey, record.parentId, statuses.get(record.parentId), false)
                );
            } else {
                if (
                    data.form.kind !== "trackerStage" &&
                    record.surveyId &&
                    parentRecord.surveyId &&
                    record.surveyId !== parentRecord.surveyId
                ) {
                    addFlag(record, {
                        flag: "SURVEY MISMATCH",
                        detail: `Survey_id ${record.surveyId}, but its ${parentKey} ${parentRecord.id} has ${parentRecord.surveyId}`,
                    });
                }
                if (parentRecord.flags.length > 0) {
                    const kinds = [...new Set(parentRecord.flags.map(f => f.flag))].join(", ");
                    addFlag(record, {
                        flag: "PARENT FLAGGED",
                        detail: `${parentKey} ${parentRecord.id} is flagged: ${kinds}`,
                    });
                }
            }
        }
    }
}

/** Flagged-record counts per sheet and flag, for _index and the console. */
export function summariseFlags(
    all: FormData[]
): { sheet: string; flag: FlagKind; count: number }[] {
    const counts = new Map<string, { sheet: string; flag: FlagKind; count: number }>();
    for (const data of all) {
        for (const record of data.records) {
            for (const flag of new Set(record.flags.map(f => f.flag))) {
                const key = `${data.form.key}|${flag}`;
                const row = counts.get(key) ?? { sheet: data.form.key, flag, count: 0 };
                row.count++;
                counts.set(key, row);
            }
        }
    }
    return [...counts.values()];
}

/** Ids per tracker lookup request. */
const LOOKUP_BATCH = 50;

/**
 * Asks DHIS2 about ids records point at but the extract does not contain, deleted records
 * included, so the flags can say "deleted" rather than just "missing".
 */
async function lookUpReferences(
    api: D2Api,
    session: SessionManager,
    ids: { surveys: string[]; trackedEntities: string[] }
): Promise<Map<string, ReferenceStatus>> {
    type Row = { deleted?: boolean; orgUnit?: string; updatedAt?: string };
    const found = new Map<string, Row>();

    const query = async (
        path: string,
        idParam: string,
        idField: "event" | "trackedEntity",
        list: string[]
    ) => {
        for (let i = 0; i < list.length; i += LOOKUP_BATCH) {
            const batch = list.slice(i, i + LOOKUP_BATCH);
            const resp = await session.retryWithBackoff(() =>
                api
                    .get<{
                        [key: string]:
                            | (Row & { event?: string; trackedEntity?: string })[]
                            | undefined;
                    }>(path, {
                        [idParam]: batch.join(","),
                        includeDeleted: true,
                        ouMode: "ACCESSIBLE",
                        fields: `${idField},deleted,orgUnit,updatedAt`,
                        pageSize: batch.length,
                    })
                    .getData()
            );
            for (const row of resp[idParam] ?? resp.instances ?? []) {
                const id = row[idField];
                if (id) found.set(id, row);
            }
        }
    };
    await query("/tracker/events", "events", "event", ids.surveys);
    await query(
        "/tracker/trackedEntities",
        "trackedEntities",
        "trackedEntity",
        ids.trackedEntities
    );

    const elsewhere = [...found.values()]
        .filter(r => !r.deleted && r.orgUnit)
        .map(r => r.orgUnit as string);
    const names = await fetchOrgUnitNames(api, session, [...new Set(elsewhere)]);

    const statuses = new Map<string, ReferenceStatus>();
    for (const [id, row] of found) {
        statuses.set(
            id,
            row.deleted
                ? { status: "deleted", lastUpdated: row.updatedAt ?? "" }
                : {
                      status: "elsewhere",
                      orgUnit: names.get(row.orgUnit ?? "") ?? row.orgUnit ?? "?",
                  }
        );
    }
    return statuses;
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
    recordIndex: Map<string, Map<string, Record_>>
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

/** Beyond 15 significant digits Excel silently rounds a number, so such values stay text. */
const EXCEL_MAX_DIGITS = 15;

/**
 * A stored value as an Excel cell: a real number or date for numeric/date fields, so they
 * can be summed and filtered, and the text unchanged for everything else (UIDs and coded
 * values must never be re-parsed). A value that does not convert exactly stays text.
 */
export function toCellValue(raw: string, valueType: string | undefined): string | number | Date {
    if (raw === "" || !valueType) return raw;

    if (NUMERIC_VALUE_TYPES.has(valueType)) {
        const n = Number(raw);
        const digits = raw.replace(/[^0-9]/g, "").replace(/^0+/, "").length;
        return Number.isFinite(n) && digits <= EXCEL_MAX_DIGITS ? n : raw;
    }
    if (valueType === "DATE" || valueType === "DATETIME") {
        // DHIS2 dates carry no zone. Read them as UTC, which is how Excel serials are
        // written, so the cell shows exactly the stored date/time.
        const hasZone = /(Z|[+-]\d\d:?\d\d)$/i.test(raw);
        const iso =
            valueType === "DATE" ? `${raw.slice(0, 10)}T00:00:00Z` : hasZone ? raw : `${raw}Z`;
        const date = new Date(iso);
        return Number.isNaN(date.getTime()) ? raw : date;
    }
    return raw;
}

function columnNumFmt(valueType: string | undefined): string {
    if (valueType === "DATE") return "yyyy-mm-dd";
    if (valueType === "DATETIME") return "yyyy-mm-dd hh:mm";
    return valueType && NUMERIC_VALUE_TYPES.has(valueType) ? "General" : "@";
}

/** Meta columns holding DHIS2 timestamps, written as real date-times. */
const META_DATETIME_KEYS = new Set([
    "created_at",
    "updated_at",
    "enrolled_at",
    "occurred_at",
    "last_updated",
]);

/** Light orange: flagged rows stand out without hiding their content. */
const FLAG_FILL: Excel.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFD8A8" } };

type Writer = Excel.stream.xlsx.WorkbookWriter;

/**
 * Streams one form sheet. Flag columns come first, so a flagged row is the first thing an
 * epidemiologist sees on it; flagged rows are also shaded, and the header has filters.
 */
export function writeFormSheet(
    workbook: Writer,
    data: FormData,
    byKey: Map<string, FormData>,
    recordIndex: Map<string, Map<string, Record_>>,
    orgUnitNames: Map<string, string>,
    sheetName: string
): void {
    const { form, records } = data;
    const columns = uniqueLabels(data.columns);
    const columnUids = [...columns.keys()];
    const keyColumns = keyColumnsFor(form, byKey);
    const metaKeys = [...new Set(records.flatMap(r => [...r.meta.keys()]))];

    const leading = [
        "flag",
        "flag_detail",
        ...keyColumns.map(c => c.header),
        "path",
        "record_id",
        "org_unit_id",
        "org_unit_name",
    ];
    const header = [...leading, ...metaKeys, ...columns.values()];
    // One valueType per column: identifiers and codes stay text; timestamps and typed fields don't.
    const types = [
        ...leading.map(() => undefined),
        ...metaKeys.map(k => (META_DATETIME_KEYS.has(k) ? "DATETIME" : undefined)),
        ...columnUids.map(uid => data.valueTypes?.get(uid)),
    ];

    const sheet = workbook.addWorksheet(sheetName, { views: [{ state: "frozen", ySplit: 1 }] });
    // Streaming writes columns with the first row, so formats are set first, at column level.
    // Text by default, so DHIS2 UIDs and coded values are never re-parsed by Excel as numbers
    // or dates if the sheet is edited.
    sheet.columns = types.map(type => ({ width: 18, style: { numFmt: columnNumFmt(type) } }));
    sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: header.length } };
    const headerRow = sheet.addRow(header);
    headerRow.font = { bold: true };
    headerRow.commit();

    if (records.length + 1 > EXCEL_MAX_ROWS) {
        console.warn(
            `  ! ${form.key}: ${records.length} rows exceeds Excel's sheet limit; only the first ` +
                `${EXCEL_MAX_ROWS - 1} are written. Narrow the scope (e.g. a smaller --org-unit) ` +
                `to extract the rest.`
        );
    }

    for (const record of records.slice(0, EXCEL_MAX_ROWS - 1)) {
        const values = [
            [...new Set(record.flags.map(f => f.flag))].join("; "),
            record.flags.map(f => f.detail).join("; "),
            ...keyColumns.map(c => c.get(record)),
            buildBreadcrumb(form, record, byKey, recordIndex),
            record.id,
            record.orgUnit,
            orgUnitNames.get(record.orgUnit) ?? "",
            ...metaKeys.map(k => record.meta.get(k) ?? ""),
            ...columnUids.map(uid => record.values.get(uid) ?? ""),
        ];
        // An empty value writes no cell at all: a truly blank cell, and a smaller file.
        const row = sheet.addRow(
            values.map((value, i) => (value === "" ? null : toCellValue(value, types[i])))
        );
        if (record.flags.length > 0) {
            for (let column = 1; column <= header.length; column++)
                row.getCell(column).fill = FLAG_FILL;
        }
        row.commit();
    }
    sheet.commit();
}

/**
 * Machine-readable relationship contract: one row per sheet giving its primary key, parent
 * key, ancestor keys, expected cardinality and an example join. Downstream tools (Power BI,
 * SQL loaders) can read this sheet instead of hard-coding the model.
 */
export function writeRelationshipsSheet(
    workbook: Writer,
    all: FormData[],
    facilityKey: string | undefined
): void {
    const sheet = workbook.addWorksheet("_relationships");
    [20, 14, 16, 16, 42, 34, 24, 52].forEach((width, i) => (sheet.getColumn(i + 1).width = width));
    const byKey = new Map(all.map(d => [d.form.key, d]));

    sheet.addRow(["Relationship specification"]).font = { bold: true, size: 14 };
    sheet.addRow([
        "Note",
        "Case report has NO foreign key to Facility: it links to Survey. Facility_id is DERIVED " +
            "by matching (Survey_id, org_unit_id) against the Facility sheet.",
    ]);
    sheet.addRow([]);

    sheet.addRow([
        "sheet",
        "primary_key",
        "parent_sheet",
        "parent_key",
        "ancestor_keys",
        "cardinality",
        "source_field",
        "example_join",
    ]).font = { bold: true };

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
            cardinality =
                "aggregate: 1 row per data element x disaggregation x period x org unit x ward";
            exampleJoin = facilityKey
                ? `${form.key}.org_unit_id = ${facilityKey}.org_unit_id`
                : "org_unit_id = <facility org unit>";
            sourceField = "Ward = the dataValue's attribute option combo";
        } else if (isSurvey) {
            parentKeyCol = "(root)";
            cardinality = "root";
            exampleJoin = "-";
            sourceField = "-";
        } else if (parentForm) {
            const isFacility = form.defaultUid === PREVALENCE_FACILITY_LEVEL_FORM_ID;
            parentKeyCol =
                isFacility || parentForm.defaultUid === PREVALENCE_SURVEY_FORM_ID
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
    sheet.addRow(["Derived key", "Rule"]).font = { bold: true };
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
    sheet.commit();
}

export function writeIndexSheet(
    workbook: Writer,
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
    [34, 62, 14, 14, 14, 14, 14, 14, 14, 14, 14].forEach(
        (width, i) => (sheet.getColumn(i + 1).width = width)
    );
    const byKey = new Map(all.map(d => [d.form.key, d]));
    const heading = (cells: (string | number)[]) => (sheet.addRow(cells).font = { bold: true });

    sheet.addRow(["AMR Surveys extraction"]).font = { bold: true, size: 14 };
    sheet.addRow(["Server", baseUrl]);
    sheet.addRow(["Extracted at", new Date().toISOString()]);
    if (extra?.country) {
        const c = extra.country;
        sheet.addRow([
            "Country",
            `${c.name}${c.code ? ` (${c.code})` : ""} uid=${c.id} level=${c.level}`,
        ]);
    }
    sheet.addRow([
        "Note",
        "Read-only extraction. One sheet per form; join on the *_id columns. See _relationships.",
    ]);
    sheet.addRow([
        "Flags",
        "Rows needing a second look are shaded orange and say why in the 'flag' and 'flag_detail' " +
            "columns (columns A-B of every sheet); filter column A to include or exclude them.",
    ]);
    sheet.addRow([]);

    heading([
        "Sheet",
        "Form name (server)",
        "UID",
        "Kind",
        "Custom",
        "Parent sheet",
        "Parent link field",
        "Records",
        "Flagged",
        "Expected",
        "Reconciled",
    ]);

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
            records.filter(r => r.flags.length > 0).length,
            expected ?? "n/a",
            reconciled,
        ]);
    }

    sheet.addRow([]);
    heading(["Hierarchy (child counts)"]);
    const renderTree = (data: FormData, depth: number) => {
        sheet.addRow([
            `${"    ".repeat(depth)}${depth > 0 ? "└─ " : ""}${data.form.key}`,
            data.records.length,
        ]);
        for (const child of all.filter(d => d.form.parentKey === data.form.key))
            renderTree(child, depth + 1);
    };
    for (const root of all.filter(d => !d.form.parentKey)) renderTree(root, 0);
    for (const detached of all.filter(d => d.form.parentKey && !byKey.has(d.form.parentKey))) {
        renderTree(detached, 0);
    }

    if (unresolved.length > 0) {
        sheet.addRow([]);
        heading(["Not extracted", "Reason"]);
        for (const { spec, reason } of unresolved) sheet.addRow([spec.requestedName, reason]);
    }

    const flagged = summariseFlags(all);
    sheet.addRow([]);
    heading(["Flagged records", "Flag", "Records"]);
    if (flagged.length === 0) sheet.addRow(["none"]);
    for (const f of flagged) sheet.addRow([f.sheet, f.flag, f.count]);
    sheet.addRow([]);
    heading(["Flag", "Meaning"]);
    for (const [flag, meaning] of Object.entries(FLAG_MEANINGS)) sheet.addRow([flag, meaning]);

    const integrity = extra?.integrity;
    if (integrity) {
        if (integrity.offFormFields.length > 0) {
            // A field can be off-form in two programs merged into one sheet: one row per column.
            const byColumn = new Map<string, { sheet: string; column: string; values: number }>();
            for (const f of integrity.offFormFields) {
                const key = `${f.sheet}|${f.column}`;
                const row = byColumn.get(key) ?? { ...f, values: 0 };
                row.values += f.values;
                byColumn.set(key, row);
            }
            sheet.addRow([]);
            heading([
                "Fields no longer on the form",
                "Column (marked [not on current form])",
                "Values",
            ]);
            sheet.addRow([
                "",
                "Data entered in fields since removed from the form. The app no longer shows them; " +
                    "this extract keeps them. Headers use the field's full DHIS2 name.",
            ]);
            const sheetOrder = new Map(all.map((d, i) => [d.form.key, i]));
            const rows = [...byColumn.values()].sort(
                (a, b) =>
                    (sheetOrder.get(a.sheet) ?? 0) - (sheetOrder.get(b.sheet) ?? 0) ||
                    a.column.localeCompare(b.column, undefined, { numeric: true })
            );
            for (const f of rows) sheet.addRow([f.sheet, f.column, f.values]);
        }

        const findings: [string, string][] = [];
        for (const o of integrity.otherDataSetValues) {
            findings.push([
                `other dataSet's values: ${o.formKey}`,
                `${o.count} value(s) returned for this dataSet belong to another dataSet's wards (the ward ` +
                    `dataSets share data elements); they are on that dataSet's sheet, not this one`,
            ]);
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
                `survey ${a.surveyId} + orgUnit ${a.orgUnit} matches ${
                    a.facilityIds.length
                }: ${a.facilityIds.join(", ")}`,
            ]);
        }
        for (const u of integrity.unresolvedStageEvents) {
            findings.push([
                `SKIPPED events: ${u.formKey}`,
                `${u.count} event(s) referenced unresolvable program stage ${u.programStage} — their data was NOT extracted`,
            ]);
        }
        for (const d of integrity.duplicateStageEvents) {
            findings.push([
                `duplicate stage events: ${d.formKey}`,
                `${d.count} extra event(s) on single-entry stage ${d.programStage} — the last event's values overwrote the earlier ones in the shared row`,
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
                `deleted field: ${h.formKey}`,
                `${h.kind} ${h.id} holds values but no longer exists in DHIS2 metadata; the header shows its id`,
            ]);
        }

        sheet.addRow([]);
        heading(["Other findings", "Detail"]);
        if (findings.length === 0) sheet.addRow(["none"]);
        for (const f of findings) sheet.addRow(f);
    }
    sheet.commit();
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
const COUNTRY_FIELDS = { id: true, name: true, code: true, level: true } as const;

/**
 * Every record traces to a Survey, and a Survey's org unit is its country, so the Survey
 * form's org units are the countries that have data.
 */
async function countriesWithData(
    api: D2Api,
    session: SessionManager,
    forms: ResolvedForm[]
): Promise<ResolvedCountry[]> {
    const survey = forms.find(f => f.defaultUid === PREVALENCE_SURVEY_FORM_ID);
    if (!survey)
        throw new Error("--per-country needs the Survey form, which could not be resolved.");

    const meta = await fetchProgramMeta(api, session, survey.uid);
    const opts: FetchOpts = { pageSize: 500, requests: createLimiter(REQUEST_CONCURRENCY) };
    const { records } = await fetchEvents(api, survey, meta, opts, session, emptyIntegrityReport());
    const ids = [...new Set(records.map(r => r.orgUnit).filter(Boolean))];
    if (ids.length === 0)
        throw new Error("--per-country found no Survey records to take countries from.");

    const { organisationUnits } = await session.retryWithBackoff(() =>
        api.metadata
            .get({ organisationUnits: { fields: COUNTRY_FIELDS, filter: { id: { in: ids } } } })
            .getData()
    );
    return organisationUnits
        .map(ou => ({ ...ou, matchedBy: "survey" }))
        .sort((a, b) => a.name.localeCompare(b.name));
}

/**
 * Resolves --country by code (e.g. KEN), then exact name, then UID. Fails loudly and lists
 * candidates on ambiguity rather than silently picking one.
 */
export async function resolveCountry(
    api: D2Api,
    session: SessionManager,
    input: string
): Promise<ResolvedCountry> {
    const term = input.trim();

    const attempts = [
        { matchedBy: "code", filter: { code: { eq: term.toUpperCase() } } },
        { matchedBy: "name", filter: { name: { eq: term } } },
        ...(UID_RE.test(term) ? [{ matchedBy: "uid", filter: { id: { eq: term } } }] : []),
    ];

    for (const attempt of attempts) {
        const { organisationUnits } = await session.retryWithBackoff(() =>
            api.metadata
                .get({ organisationUnits: { fields: COUNTRY_FIELDS, filter: attempt.filter } })
                .getData()
        );

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
    return (
        keyColumnsFor(form, byKey).length + FIXED_NON_META_COLUMNS + metaKeyCount + valueColumnCount
    );
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
            fetchProgramMeta(api, session, form.uid),
        ]);
        return {
            form,
            total: first.total ?? 0,
            minDate: first.instances[0]?.occurredAt,
            maxDate: last.instances[0]?.occurredAt,
            columns: computeColumnCount(
                form,
                byKey,
                EVENT_META_KEYS.length,
                [...meta.stageById.values()].reduce((sum, s) => sum + s.dataElementLabels.size, 0)
            ),
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

    const [first, last, meta] = await Promise.all([
        probe("asc"),
        probe("desc"),
        fetchProgramMeta(api, session, form.uid),
    ]);

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

    return mapWithConcurrency(forms, FETCH_CONCURRENCY, f =>
        discoverForm(api, f, orgUnit, session, byKey)
    );
}

/** "CaseReport [custom]" / "CaseReport [default]": which program of a form is meant. */
function programLabel(form: ResolvedForm): string {
    return `${form.key} [${form.isCustom ? "custom" : "default"}]`;
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
            warnings.push(
                `${f.sheetKey}: ${f.columns} columns — wide sheet, slow to browse by hand in Excel`
            );
        }
        if (f.rows !== undefined && f.rows > LARGE_SHEET_ROW_WARNING) {
            warnings.push(
                `${
                    f.sheetKey
                }: ${f.rows.toLocaleString()} rows — approaching Excel's 1,048,576-row limit`
            );
        }
    };

    for (const d of discoveries.filter(d => d.total !== 0)) {
        addSheet({
            sheetKey: d.form.key,
            rows: d.total,
            columns: d.columns,
            cells:
                d.total !== undefined && d.columns !== undefined ? d.total * d.columns : undefined,
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
        estimatedSizeBytesRange: [
            totalCells * BYTES_PER_CELL_LOW,
            totalCells * BYTES_PER_CELL_HIGH,
        ],
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

    console.log("\nForms detected (one row per program; programs of one form share a sheet):");
    console.log(
        "  " +
            [
                "SHEET".padEnd(18),
                "PROGRAM".padEnd(13),
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
                    d.form.uid.padEnd(13),
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
    console.log(`  programs with data   : ${withData.length}/${discoveries.length}`);
    console.log(`  main-sheet records   : ${totalRecords}`);
    if (totalStageRecords > 0) {
        console.log(`  repeatable-stage rows: ${totalStageRecords}`);
    }
    if (empty.length > 0) {
        console.log(`  skipped (0 records): ${empty.map(d => programLabel(d.form)).join(", ")}`);
    }
    const dataSets = discoveries.filter(d => d.form.kind === "dataSet");
    if (dataSets.length > 0) {
        console.log(
            `  ${dataSets
                .map(d => d.form.key)
                .join(
                    ", "
                )}: aggregate dataSet — all periods unless --start-date/--end-date are given`
        );
    }
    if (unresolved.length > 0) {
        console.log(`  not extracted: ${unresolved.map(u => u.spec.requestedName).join("; ")}`);
    }

    const forecast = forecastWorkbook(discoveries);
    console.log(`\nWorkbook forecast:`);
    const formSheets = new Set(forecast.sheets.map(s => s.sheetKey)).size;
    console.log(
        `  worksheets    : ${
            2 + formSheets
        }  (_index, _relationships + ${formSheets} form sheet(s))`
    );
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

type ExtractArgs = {
    output?: string;
    country?: string;
    orgUnit?: string;
    startDate?: string;
    endDate?: string;
    pageSize: number;
    dryRun: boolean;
    discover: boolean;
    perCountry: boolean;
    forms?: string;
};

/** Everything one scope's extraction needs that does not depend on the scope. */
type RunContext = {
    api: D2Api;
    session: SessionManager;
    baseUrl: string;
    envLabel: string;
    resolved: ResolvedForm[];
    unresolved: { spec: FormSpec; reason: string }[];
    args: ExtractArgs;
};

async function extract(args: ExtractArgs) {
    const envVars = getEnvVars();
    const baseUrl = envVars.url.replace(/\/+$/, "");
    const envLabel = deriveEnvLabel(baseUrl);

    console.log(`Connecting to ${baseUrl} (${describeAuth(envVars)}) [env: ${envLabel}]`);
    const api = getD2APiFromInstance(getInstance(envVars));

    // Forces DHIS2 to fully initialize the session before any real call (PAT bug workaround).
    await warmUpSession(api);
    const session = createSessionManager(api);

    const info = await session.retryWithBackoff(() => api.system.info.getData());
    console.log(`  DHIS2 ${info.version}`);

    const modules = await fetchModules(api, session);
    console.log(`  datastore modules: ${modules.map(m => m.name).join(", ") || "(none)"}`);

    const wanted = args.forms
        ? FORMS.filter(f =>
              args
                  .forms!.split(",")
                  .map(s => s.trim())
                  .includes(f.key)
          )
        : FORMS;
    if (wanted.length === 0) throw new Error(`No forms matched --forms=${args.forms}`);

    const { resolved, unresolved } = await resolveForms(api, session, modules, wanted);
    linkParents(resolved);
    reportResolution(resolved, unresolved);

    if (args.dryRun) {
        console.log("\nDry run: no data extracted, no file written.");
        return;
    }
    if (resolved.length === 0) throw new Error("No forms could be resolved; nothing to extract.");

    const ctx: RunContext = { api, session, baseUrl, envLabel, resolved, unresolved, args };

    if (args.perCountry) {
        if (args.country || args.orgUnit) {
            throw new Error("--per-country cannot be combined with --country or --org-unit.");
        }
        const countries = await countriesWithData(api, session, resolved);
        console.log(
            `\nCountries with data (${countries.length}): ${countries
                .map(c => c.code || c.name)
                .join(", ")}`
        );

        // One country failing (after its retries) must not cost the others their workbooks.
        const results: { country: ResolvedCountry; programs?: string[]; error?: string }[] = [];
        for (const country of countries) {
            console.log(`\n=== ${country.name}${country.code ? ` (${country.code})` : ""} ===`);
            try {
                results.push({ country, programs: await extractScope(ctx, country, country.id) });
            } catch (err) {
                const error = err instanceof Error ? err.message : String(err);
                console.error(`  ! ${country.name}: FAILED — ${error}`);
                results.push({ country, error });
            }
        }

        console.log("\nPrograms used per country:");
        for (const { country, programs, error } of results) {
            const label = country.code || country.name;
            console.log(
                `  ${label}: ${error ? `FAILED — ${error}` : programs?.join("; ") || "(none)"}`
            );
        }
        const failed = results.filter(r => r.error);
        if (failed.length > 0) {
            throw new Error(
                `${failed.length} of ${results.length} countries failed: ${failed
                    .map(r => r.country.code || r.country.name)
                    .join(", ")}`
            );
        }
        return;
    }

    const country = args.country ? await resolveCountry(api, session, args.country) : undefined;
    await extractScope(ctx, country, args.orgUnit ?? country?.id);
}

/**
 * Discovers, extracts and writes one workbook for one scope (a country, an org unit, or
 * everything readable). Returns the programs it extracted, as "Form: program (records)".
 */
async function extractScope(
    ctx: RunContext,
    country: ResolvedCountry | undefined,
    rootOrgUnit: string | undefined
): Promise<string[]> {
    const { api, session, baseUrl, envLabel, resolved, unresolved, args } = ctx;
    const started = Date.now();
    const elapsed = () => `${((Date.now() - started) / 1000).toFixed(1)}s`;

    // --- Discovery ------------------------------------------------------------
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
        return [];
    }

    // Only skip programs proven to have exactly zero records — never narrow on anything else.
    const totalByUid = new Map(discoveries.map(d => [d.form.uid, d.total]));
    const toExtract = resolved.filter(f => totalByUid.get(f.uid) !== 0);
    const skipped = resolved.filter(f => totalByUid.get(f.uid) === 0);
    if (skipped.length > 0) {
        console.log(
            `\nSkipping ${skipped.length} program(s) with 0 records: ${skipped
                .map(programLabel)
                .join(", ")}`
        );
    }

    console.log(`\nExtracting ... (discovery took ${elapsed()})`);
    const opts: FetchOpts = {
        orgUnit: rootOrgUnit,
        pageSize: args.pageSize,
        startDate: args.startDate,
        endDate: args.endDate,
        requests: createLimiter(REQUEST_CONCURRENCY),
        // Without a scope, dataSets are read under the user's own root org units.
        dataSetRoots:
            !rootOrgUnit && toExtract.some(f => f.kind === "dataSet")
                ? await userRootOrgUnits(api, session)
                : undefined,
    };

    // Declared before extraction (not after) so fetchTracker can record diagnostics
    // (e.g. unresolved stage events) as they happen, not as an afterthought.
    const integrity = emptyIntegrityReport();

    // Programs are independent, so several are fetched at once; their page requests share
    // one cap (opts.requests). `mapWithConcurrency` preserves input order, so sheet order
    // stays deterministic regardless of which program finishes first.
    const perProgram = (
        await mapWithConcurrency(toExtract, EXTRACT_CONCURRENCY, async form => {
            const formOpts = { ...opts, total: totalByUid.get(form.uid) };
            if (form.kind === "dataSet")
                return [await fetchDataSet(api, form, formOpts, session, integrity)];
            const meta = await fetchProgramMeta(api, session, form.uid);
            if (form.kind !== "tracker")
                return [await fetchEvents(api, form, meta, formOpts, session, integrity)];
            const { main, stageForms } = await fetchTracker(
                api,
                form,
                meta,
                formOpts,
                session,
                integrity
            );
            for (const stage of stageForms) {
                console.log(`  ${stage.form.key}: ${stage.records.length} rows (repeatable stage)`);
            }
            return [main, ...stageForms];
        })
    ).flat();
    console.log(`  fetched in ${elapsed()}`);

    const programs = perProgram
        .filter(d => d.form.kind !== "trackerStage")
        .map(d => `${d.form.key}: ${d.form.serverName} (${d.records.length})`);

    // One sheet per form, however many programs back it.
    const all = mergeByFormKey(perProgram);

    const byKey = new Map(all.map(d => [d.form.key, d]));
    const recordIndex = new Map(all.map(d => [d.form.key, new Map(d.records.map(r => [r.id, r]))]));

    // --- Integrity ------------------------------------------------------------
    const facilityForm = resolved.find(f => f.defaultUid === PREVALENCE_FACILITY_LEVEL_FORM_ID);
    if (facilityForm) resolveFacilityIds(all, facilityForm.key, integrity);
    await flagRecords(all, recordIndex, ids => lookUpReferences(api, session, ids));

    // Reconciliation: what discovery said exists vs what we actually pulled. This is what
    // proves the skip/scope optimisation above never silently dropped data.
    const expected = expectedRecordCounts(toExtract, totalByUid, discoveries);
    const reconciliation = all.map(d => ({
        formKey: d.form.key,
        expected: expected.get(d.form.key),
        actual: d.records.length,
    }));
    const mismatches = reconciliation.filter(
        r => r.expected !== undefined && r.expected !== r.actual
    );
    if (mismatches.length > 0) {
        console.warn("\n  ! RECONCILIATION MISMATCH — extracted count != discovered count:");
        for (const m of mismatches) {
            console.warn(`      ${m.formKey}: expected ${m.expected}, got ${m.actual}`);
        }
    } else if (discoveries.length > 0) {
        console.log("\nReconciliation: all forms match their discovered counts.");
    }

    const orgUnitIds = [
        ...new Set(all.flatMap(d => d.records.map(r => r.orgUnit)).filter(Boolean)),
    ];
    console.log(`\nResolving ${orgUnitIds.length} org unit names ...`);
    const orgUnitNames = await fetchOrgUnitNames(api, session, orgUnitIds);

    // Name the file after the instance and country so an extract is never ambiguous
    // about where it came from (mirrors glass-dev's deriveEnvLabel convention).
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    const countryPart = country ? `${(country.code || country.name).toLowerCase()}_` : "";
    const fileName = `amr-surveys_${envLabel}_${countryPart}${timestamp}.xlsx`;

    // With --per-country, --output names a directory (one file per country goes in it).
    const outPath = args.perCountry
        ? path.resolve(args.output ?? "extracts", fileName)
        : path.resolve(args.output ?? `extracts/${fileName}`);
    // Written under a temporary name and renamed when complete, so a run that dies midway
    // never leaves a truncated workbook that looks like a finished extract.
    const partialPath = `${outPath}.partial`;
    fs.mkdirSync(path.dirname(outPath), { recursive: true });

    console.log("Writing workbook ...");
    // Streaming: rows go to disk as they are written instead of building the whole workbook
    // in memory first, which was two thirds of a country's run time.
    // Its zip library defaults to the fastest compression (level 1); level 6 makes the file
    // about a quarter smaller for a second or two of CPU.
    const workbook = new Excel.stream.xlsx.WorkbookWriter({
        filename: partialPath,
        useStyles: true,
        useSharedStrings: true,
        zip: { zlib: { level: 6 } },
    });
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
        writeFormSheet(workbook, data, byKey, recordIndex, orgUnitNames, name);
    }
    await workbook.commit();
    fs.renameSync(partialPath, outPath);

    const total = all.reduce((sum, d) => sum + d.records.length, 0);
    console.log(
        `\nDone in ${elapsed()}. ${total} records across ${all.length} sheets -> ${outPath}`
    );

    const facilities = facilityForm ? facilitiesWithData(all, facilityForm.key) : new Set();
    if (facilities.size > 0) console.log(`  facilities with data: ${facilities.size}`);
    reportIntegrity(integrity, all);
    return programs;
}

/** The user's root org units for data reads: data-view roots, else capture roots. */
async function userRootOrgUnits(api: D2Api, session: SessionManager): Promise<string[]> {
    const me = await session.retryWithBackoff(() =>
        api
            .get<{
                organisationUnits?: { id: string }[];
                dataViewOrganisationUnits?: { id: string }[];
            }>("/me", {
                fields: "organisationUnits[id],dataViewOrganisationUnits[id]",
            })
            .getData()
    );
    const roots = me.dataViewOrganisationUnits?.length
        ? me.dataViewOrganisationUnits
        : me.organisationUnits ?? [];
    return roots.map(ou => ou.id);
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
                    "Output .xlsx path (a directory with --per-country). Defaults to extracts/amr-surveys_<env>_<country>_<timestamp>.xlsx",
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
                description:
                    "Start date YYYY-MM-DD (Ward Summary Statistics only; default 2000-01-01)",
            }),
            endDate: option({
                type: optional(string),
                long: "end-date",
                description: "End date YYYY-MM-DD (Ward Summary Statistics only; default today)",
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
            perCountry: flag({
                type: boolean,
                long: "per-country",
                description:
                    "Write one workbook per country that has data (countries are taken from the Survey records), and print the programs each country used.",
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
