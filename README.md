# AMR Surveys
AMR Surveys is an app to manage data input and configuration of surveys for AMR.

## User guide

### Hide elements or sections by survey
1. Based on the rules defined in datastore for each parent survey instance, apply HIDEFIELD/HIDESECTION rule for each child form.
2. refactor code in src/domain/usecases/GetSurveyUseCase.ts to break up into functions.

NOTE : datastore structure for survey rules
 ```
"rulesBySurvey": [
      {
        "surveyId": "p91asa2vebZ",
        "surveyRules": [
          {
            "formId": "mesnCzaLc7u",
            "rules": [
              {
                "id": "1",
                "toHide": [
                  "SCYImStXDHM"
                ],
                "type": "HIDEFIELD"
              },
              {
                "id": "2",
                "toHide": [
                  "mIHOCwjHtyu"
                ],
                "type": "HIDESECTION"
              }
            ]
          }
        ]
      }
    ],

```
### :video_camera: Screenshots/Screen capture

[video](https://github.com/EyeSeeTea/amr-surveys/assets/83749675/1c200566-dabf-4e43-aa9a-b40910448c93)

# Developer guide

## Setup

```
$ nvm use # uses node version in .nvmrc
$ yarn install
```

## Build

Build a production distributable DHIS2 zip file:

```
$ yarn build
```

## Development

Copy `.env` to `.env.local` and configure DHIS2 instance to use. Then start the development server:

```
$ yarn start
```

Now in your browser, go to `http://localhost:8081`.

## Tests

```
$ yarn test
```

## Smoke test

Minimum check to run after every code change, before committing or opening a PR.
It catches runtime regressions that the type-checker and unit tests do not.

1. **Type-check**: `yarn tsc --noEmit` — must pass clean.
2. **Lint**: `yarn lint` — must pass clean.
3. **Boot the dev server**: `yarn start` and wait for vite to report `ready in …`.
4. **Browse the app** at `http://localhost:8081` (a headless browser such as
   Playwright works well for automating this):
    - Log in (DHIS2 basic-auth proxy uses the credentials in `.env.local`).
    - Open the **Prevalence** module → pick a country survey → drill into a
      facility → open the patients list (Case Reports) → open one case report
      form.
    - Open the **PPS** module → pick a country survey → drill into a hospital →
      open the patients list.
    - Open **Ward Summary Statistics**.
5. **Pass criteria**: zero uncaught exceptions in the browser console
   (`error`-level messages), no `Failed to fetch` toasts, every page reaches a
   rendered state (no infinite spinners).

If any step fails, the change is not ready to commit. The smoke test does
**not** replace unit tests or manual QA against real data — it is the floor,
not the ceiling.

## Some development tips

### Clean architecture folder structure

-   `src/domain`: Domain layer of the app (entities, use cases, repository definitions)
-   `src/data`: Data of the app (repository implementations)
-   `src/webapp/pages`: Main React components.
-   `src/webapp/components`: React components.
-   `src/utils`: Misc utilities.
-   `i18n/`: Contains literal translations (gettext format)
-   `public/`: General non-React webapp resources.

## Data structures

-   `Future.ts`: Async values, similar to promises, but cancellables and with type-safe errors.
-   `Collection.ts`: Similar to Lodash, provides a wrapper over JS arrays.
-   `Obj.ts`: Similar to Lodash, provides a wrapper over JS objects.
-   `HashMap.ts`: Similar to ES6 map, but immutable.
-   `Struct.ts`: Base class for typical classes with attributes. Features: create, update.
-   `Either.ts`: Either a success value or an error.

## Docs

We use [TypeDoc](https://typedoc.org/example/):

```
$ yarn generate-docs
```

### i18n

Update i18n .po files from `i18n.t(...)` calls in the source code:

```
$ yarn localize
```

### Scripts

Check the example script, entry `"script-example"`in `package.json`->scripts and `src/scripts/example.ts`.

#### Extracting form data (`extract-forms`)

`src/scripts/extract-forms.ts` exports the Prevalence forms into a single Excel
workbook: one sheet per form, plus an `_index` sheet (scope, counts, reconciliation,
integrity findings) and a `_relationships` sheet (the machine-readable join contract).
It is **read-only** — it issues GET requests only and never writes to DHIS2.

##### Credentials

Uses the **same env vars as the glass-dev scripts**, so one `.env.local` drives both.
`.env.local` is git-ignored and is *not* loaded automatically — pass `DOTENV_CONFIG_PATH`:

```
$ cp .env.template .env.local     # then fill in REACT_APP_DHIS2_BASE_URL + a token or auth
$ set DOTENV_CONFIG_PATH=.env.local          # Windows cmd
$ $env:DOTENV_CONFIG_PATH=".env.local"       # PowerShell
```

Auth precedence (see `src/scripts/common.ts`): `REACT_APP_DHIS2_TOKEN_PROD` >
`_PREPROD` > `_TRAINING` > `REACT_APP_DHIS2_TOKEN` > `REACT_APP_DHIS2_AUTH`
(`username:password`). Personal Access Tokens bypass 2FA and are preferred.

##### Usage

```
$ yarn extract-forms --country KEN --discover      # preflight: what exists, no extraction
$ yarn extract-forms --dry-run                     # resolve form names -> UIDs only
$ yarn extract-forms --country KEN                 # extract everything for one country
$ yarn extract-forms --country KEN --start-date 2024-01-01 --end-date 2024-12-31
```

`--country` accepts an ISO3 code (`KEN`), an exact org unit name (`Kenya`), or a UID.
`--discover` reports the resolved country, forms detected (including any repeatable-stage
child sheets, with their own row/column counts — see below), record counts, min/max dates,
a full workbook forecast (worksheet count, total cells, an estimated file size range, and
warnings for unusually wide/large sheets), and the planned scope, then exits without writing
anything. The forecast mirrors extraction's own skip rule exactly (a form or stage proven to
have 0 records produces no sheet), so the sheet list `--discover` prints is the sheet list a
real run will produce — not a superset of it. Ward Summary Statistics is an aggregate dataSet
and needs `--start-date` / `--end-date` to be extracted.

##### Relationship specification

Every sheet's primary key is `record_id`. The workbook's `_relationships` sheet emits this
same table at runtime for downstream tools.

| Sheet | Parent sheet | Parent key | Ancestor key columns | Cardinality |
| ----- | ------------ | ---------- | -------------------- | ----------- |
| `Survey` | (root) | — | — | root |
| `Facility` | Survey | `Survey_id` (`Log2Y4uqBBo`) | `Survey_id` | Survey 1→N |
| `CaseReport` | Survey | `Survey_id` (`tlRPoWumrSa`) | `Survey_id`, `Facility_id` | Survey 1→N |
| `SampleShipment` | CaseReport | `CaseReport_id` (`mUaaSzbeMmj`) | `Survey_id`, `Facility_id`, `CaseReport_id` | CaseReport 1→N |
| `CentralRefLab` | CaseReport | `CaseReport_id` (`mUaaSzbeMmj`) | idem | CaseReport 1→N |
| `PathogenIsolates` | CaseReport | `CaseReport_id` (`M1D2XXokPWl`) | idem | CaseReport 1→N |
| `Supranational` | CaseReport | `CaseReport_id` (`yq8en6ZkENB`) | idem | CaseReport 1→N |
| `FollowUpD28` | CaseReport | `CaseReport_id` (`l4Y96YlhYyF`) | idem | CaseReport 1→N |
| `DischargeClinical` | CaseReport | `CaseReport_id` (`oT6f0BG74xs`) | idem | CaseReport 1→0..N |
| `DischargeEconomic` | CaseReport | `CaseReport_id` (`HkBG3DVELBM`) | idem | CaseReport 1→0..N |
| `CohortEnrolment` | CaseReport | `CaseReport_id` (`mGYxog3at84`) | idem | CaseReport 1→0..N |
| `WardSummaryStats` | (none) | — | `org_unit_id`, `ward_form_id` | aggregate |

**Case report links to Survey, not Facility.** This is not obvious from the UI's visual
nesting, but it is what the app filters on: `GetPaginatedSurveysUseCase.ts` uses
`parentPatientId` only for the 8 leaf forms (`isPrevalencePatientChild`) and
`parentSurveyId` for Facility and Case report. So `tlRPoWumrSa` holds a **Survey** id.
Case report has no foreign key to Facility at all — the UI scopes it by org unit.

Consequently:

-   `Survey_id` is read **directly** from each record's own survey-link attribute
    (every leaf form carries one — see `parentPrevalenceSurveyIdList`), so every record
    traces to its Survey without walking the chain.
-   `Facility_id` is **derived**: `Facility.record_id WHERE Facility.Survey_id = row.Survey_id
    AND Facility.org_unit_id = row.org_unit_id`. Ambiguous or unresolvable matches are
    reported in `_index` rather than silently guessed.

Example join (SQL-ish):

```sql
SELECT f.*, c.Survey_id, fac.org_unit_name
FROM   FollowUpD28 f
JOIN   CaseReport  c   ON f.CaseReport_id = c.record_id
JOIN   Facility    fac ON f.Facility_id   = fac.record_id
JOIN   Survey      s   ON f.Survey_id     = s.record_id;
```

Sheets are flat tables with a single header row, stable snake_case keys and text-formatted
ids, so they load directly into Power BI or a SQL loader as a star schema (`Survey` and
`Facility` as dimensions, the leaf forms as facts).

##### Notes

-   Form UIDs are resolved from **live server metadata by name**, not hardcoded.
    "Custom" forms are configured per parent survey in the datastore
    (`amr-surveys` -> `modules` -> `customForms`) and so cannot be known from the
    code alone. The constants in `src/data/entities/D2Survey.ts` are used only as
    a cross-check, and a mismatch is reported as a warning.
-   Optimisation never narrows the result: the only forms skipped are those a probe proves
    have **exactly zero** records, and a reconciliation check compares extracted counts
    against the discovered totals, reporting any mismatch in `_index`.
-   `src/scripts/common.ts` is a deliberate port of glass-dev's script auth layer
    (PAT handling, `backend: "fetch"`, session warm-up, refresh + backoff). Keep the two
    in sync when either changes.

##### Program-stage data (not just attributes)

A tracker program's data lives in two places: tracked-entity **attributes**, and events on
the enrollment's **program stages**. Both are extracted:

-   A **non-repeatable** stage has at most one event per record, so its dataValues are
    flattened onto the same row as extra columns, labelled `<stage name>: <field name>`
    to disambiguate fields that share a name across stages.
-   A **repeatable** stage (e.g. Facility's "Ward data", which can occur many times per
    facility) becomes its **own child sheet** — one row per event — because a TEI can have
    many events for that stage, which is a different cardinality than the flat one-row-per-
    record model everything else uses. Its sheet key is `<owning form>__<StageName>` in
    PascalCase (e.g. `Facility__WardData` — word boundaries are preserved, not stripped, so
    "Ward data" doesn't become the unreadable `Warddata`), and it carries the same ancestor
    key columns as any other child sheet, plus `<owning form>_id` (deduplicated against an
    already-present ancestor column of the same name where the owner IS that ancestor, e.g.
    Facility's own repeatable stages don't get `Facility_id` twice).

This matters for completeness: on the AMR Surveys prevalence programs, the large majority of
real fields live at the stage level, not as attributes (e.g. Case report currently has ~17
attributes but 6 stages totalling ~195 additional fields). Extracting attributes alone would
silently drop most of the data.

##### Sheet-name safety and diagnostics

Excel sheet names must be unique and <=31 characters. `assignSheetNames()` computes every
sheet's actual tab name in one deterministic pass: names are sorted by their stable internal
key (never by API/discovery encounter order, which is not guaranteed stable run-to-run for a
live dataset), sanitized, and truncated. If two different keys would truncate to the *same*
31-char name — verified live: two `SampleShipment` stage names differing only after
character 31 — the later one gets a short, stable disambiguation suffix derived from its own
DHIS2 uid, not a running counter, so the same input metadata always produces the same output
names. `_index` and `_relationships` are reserved so a form can never collide with them.

Nothing fails silently: a resolved collision, an event on an unresolvable program stage
(data that had to be skipped), an ambiguous Facility match, and a Survey/parent id mismatch
are all counted and listed in the `_index` sheet's "Integrity findings" section, not just
logged and forgotten.

### Misc Notes

-   Requests to DHIS2 will be transparently proxied (see `vite.config.ts` -> `server.proxy`) from `http://localhost:8081/dhis2/xyz` to `${VITE_DHIS2_BASE_URL}/xyz`. This prevents CORS and cross-domain problems.

-   You can use `.env` variables within the React app: `const value = import.meta.env.NAME;`


