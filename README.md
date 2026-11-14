# Enterprise Demand Intelligence Platform (EDI)

A modular, dataset-driven Python application for discovering and loading business datasets, validating and cleaning records, preparing analytical features, generating business analysis and forecasts, producing recommendations, and creating a Power BI project. The repository also contains a multi-company data consolidation utility and Streamlit interfaces.

This README describes the source code supplied in the reviewed `EDI_Source_Code_Pack.zip`. It documents the behaviour implemented in that source pack; it is not a claim that every workflow has been end-to-end executed on every supported input format.

---

## 1. Project capabilities

The supplied code implements these main areas:

- Dataset discovery and ingestion for JSON, CSV, Excel (`.xlsx` and `.xls`), and Parquet.
- Runtime discovery of tables, columns, data types, structural keys, relationships, and temporal characteristics.
- Data validation, cleaning, standardisation, and analytical feature preparation.
- Dataset-level and demand-signal analysis.
- Demand forecasting.
- Evidence-based recommendation generation, with optional external AI enrichment in the recommendation engine.
- Run-specific pipeline workspaces, manifests, artifact validation, hashing, and cross-run stage reuse.
- Power BI PBIP/PBIR/TMDL project generation and validation.
- A Streamlit dashboard for reviewing completed runs and their outputs.
- A separate Streamlit company/workbook intelligence view with optional local Ollama integration.
- A separate company-dataset consolidation utility.
- Authenticated source-code propagation and controlled cleanup of generated run artifacts.

The project contains separate components with different responsibilities. In particular, the Data Consolidation utility is not automatically invoked by `main.py`; its workbook must be selected as the main pipeline input separately.

---

## 2. Architecture

### Main EDI pipeline

```text
Selected raw dataset
        |
        v
Dataset discovery and loading
        |
        v
Cleaning and validation
        |
        v
Preprocessing and analytical feature preparation
        |
        +--------------------------+
        |                          |
        v                          v
Business analysis             Forecasting
        |                          |
        +------------+-------------+
                     |
                     v
             Recommendation engine
                     |
                     v
             Power BI integration
                     |
                     v
        Run-specific artifacts and manifest
                     |
                     v
             Streamlit dashboard
```

The orchestrator in `main.py` defines the pipeline stages in this order:

1. `cleaning`
2. `preprocessing`
3. `analysis`
4. `forecasting`
5. `recommendation`
6. `powerbi`

Analysis and forecasting run concurrently by default. Set `EDI_PARALLEL_ANALYTICS=0` to run those two stages sequentially.

The orchestrator records run IDs, dataset fingerprints, stage status, artifact records, and a run manifest. It validates expected stage artifacts and can reuse valid artifacts from earlier runs when the cache checks pass.

### Separate Data Consolidation workflow

```text
Data_Consolidation/input/
        |
        v
Discover supported company dataset files
        |
        v
Read records and discover source schemas
        |
        v
Profile and match column headers
        |
        v
Build a combined output matrix
        |
        v
Validate and publish one Excel workbook
        |
        v
Data_Consolidation/output/
    Consolidated_Company_Dataset.xlsx
```

The consolidation utility is independent of the main pipeline execution. To analyse its resulting workbook with the main EDI pipeline, set `EDI_DATASET_PATH` to that workbook or place it as the sole supported input file in `data/raw/`.

---

## 3. Repository structure

The following is the application structure represented by the supplied source pack and the project tree provided during review. Some directories are created or populated at runtime.

```text
Enterprise_Demand_Intelligence/
│
├── config.py
├── env.py
├── main.py
├── powerbi_integration.py
├── project_cleaner.py
├── requirements.txt
├── README.md
│
├── Data_Consolidation/
│   ├── consolidate_company_datasets.py
│   ├── README.txt
│   ├── input/
│   └── output/
│
├── src/
│   ├── __init__.py
│   ├── data_loader.py
│   ├── validation.py
│   ├── cleaning.py
│   ├── data_profiler.py
│   ├── preprocessing.py
│   ├── excel_exporter.py
│   ├── analysis.py
│   ├── forecasting.py
│   └── recommendation_engine.py
│
├── streamlit_app/
│   ├── app.py
│   ├── authentication.py
│   ├── analytics.py
│   ├── component_registry.py
│   ├── data_service.py
│   ├── independent_intelligence_app.py
│   └── visualizations.py
│
├── data/
│   ├── raw/
│   ├── processed/
│   └── output/
│       └── runs/
│
├── powerbi/
│   ├── data/
│   ├── documentation/
│   └── Demand_Intelligence_Report_<RUN_ID>/
│
├── .streamlit/
├── .env_backups/
├── _patch_inspection/
└── .git/
```

`data/`, `powerbi/`, `.streamlit/`, `.env_backups/`, `_patch_inspection/`, and `.git/` may contain local configuration, generated artifacts, backups, inspection files, or version-control data. Their contents are not all part of the Python application source.

---

## 4. Python module responsibilities

### Project-root modules

| File | Responsibility |
|---|---|
| `config.py` | Central project configuration: directories, supported dataset formats, discovery and processing settings, validation controls, and output paths. |
| `env.py` | Environment/path helpers, project directory initialization, operator authentication, Python source discovery, source backups, global and targeted text replacements, syntax checks, AST comparison, and rollback. |
| `main.py` | Main pipeline orchestrator. Resolves the input dataset, creates isolated run contexts, runs pipeline stages, validates artifacts, records manifests, and optionally launches Streamlit. |
| `powerbi_integration.py` | Stages pipeline workbooks and generates, validates, and records Power BI PBIP/PBIR/TMDL project artifacts. It also supports a Power BI-only run using valid artifacts from a prior completed run. |
| `project_cleaner.py` | Inspects and cleans explicitly managed generated run artifacts with retention controls, authentication, confirmation, and protected-path verification. |

### `src/` modules

| File | Responsibility |
|---|---|
| `data_loader.py` | Validates and loads supported input files, discovers the active dataset, profiles tables and columns, and discovers structural and relationship evidence. |
| `validation.py` | Dataset structure, missing values, duplicates, key candidates, primary/foreign-key relationships, numeric values, temporal values, and type-consistency checks. |
| `cleaning.py` | Cleans and standardises discovered tables and values, performs integrity checks, and writes cleaned data and a report. |
| `data_profiler.py` | Standalone workbook/dataset profiling utilities, including structure, types, missingness, duplicates, cardinality, and candidate-key evidence. |
| `preprocessing.py` | Prepares analytical tables and dynamically derived features, validates prepared data, and writes JSON/Excel outputs and a report. |
| `excel_exporter.py` | Exports discovered tables to Excel, handles worksheet naming and Excel row limits, validates the export, and writes an export report. |
| `analysis.py` | Produces dataset overview, numeric/categorical/temporal analysis, correlations, changes, relationship evidence, and demand-signal analysis. |
| `forecasting.py` | Discovers forecastable series and temporal structure, generates forecasts, and writes forecast outputs and reports. |
| `recommendation_engine.py` | Combines analysis, signal, and forecast evidence to generate recommendations, writes workbook/report outputs, and supports optional AI enrichment through a configured HTTP endpoint with deterministic fallback logic. |

### `Data_Consolidation/`

| File | Responsibility |
|---|---|
| `consolidate_company_datasets.py` | Reads company datasets from its `input/` directory, discovers and aligns fields, preserves source provenance, and publishes a consolidated Excel workbook. |

The consolidation script accepts `.json`, `.csv`, `.xlsx`, `.xls`, and `.parquet` input files. Its output worksheet is `All_Consolidated_Data`, and the default output file is `output/Consolidated_Company_Dataset.xlsx`.

The script defines explicit markers for fields absent from a company's dataset and values not provided in an original record. It also uses semantic header matching. In the supplied source, `SEMANTIC_MODEL_REQUIRED` is set to `True` after an earlier assignment; consequently, the semantic matching dependency/model is required for a successful consolidation run.

### `streamlit_app/`

| File | Responsibility |
|---|---|
| `app.py` | Main authenticated, read-only Streamlit dashboard for pipeline runs, datasets, reports, and charts. |
| `authentication.py` | Verifies dashboard credentials using the EDI admin environment variables. |
| `analytics.py` | Dataset-adaptive metrics, schema profiling, type grouping, and nested-value expansion. |
| `component_registry.py` | Generic component-signature and field-normalisation registry. |
| `data_service.py` | Reads run manifests and exposes successful-run history and an allowlisted set of pipeline artifacts to the dashboard. |
| `independent_intelligence_app.py` | Independent Excel workbook exploration, entity profiling, adaptive charts, and evidence-grounded recommendations; it can call a local Ollama endpoint. |
| `visualizations.py` | Streamlit/Plotly chart components for numeric, categorical, boolean, missing-value, time-series, relationship, geographic, forecast, and recommendation views. |

The main dashboard also embeds the independent intelligence view. The independent workbook application can additionally be launched as its own Streamlit app.

---

## 5. Supported dataset formats and input selection

### Main pipeline

`config.py` and `src/data_loader.py` define support for:

- `.json`
- `.csv`
- `.xlsx`
- `.xls`
- `.parquet`

The main pipeline resolves the dataset in this order:

1. An explicitly configured `config.ACTIVE_DATASET_PATH`, if set.
2. The environment variable named by `config.DATASET_PATH_ENV_VAR`, which is `EDI_DATASET_PATH` in the supplied configuration.
3. Automatic discovery in `data/raw/`, if enabled.

Automatic discovery expects exactly one supported dataset file in `data/raw/`. If there are no candidates, the loader reports that no dataset was found. If there is more than one candidate, it reports an ambiguity instead of silently choosing one.

The main orchestrator also checks the supported extension and the configured maximum dataset size. The supplied configuration sets the maximum to 500 MB.

### Data Consolidation utility

Place the individual source files in:

```text
Data_Consolidation/input/
```

Run the consolidation script. Its default output is:

```text
Data_Consolidation/output/Consolidated_Company_Dataset.xlsx
```

To use that workbook as the main pipeline's dataset in PowerShell:

```powershell
$env:EDI_DATASET_PATH = (Resolve-Path ".\Data_Consolidation\output\Consolidated_Company_Dataset.xlsx").Path
```

Alternatively, copy the workbook into `data/raw/` and ensure it is the only supported dataset file there.

---

## 6. Installation

### Requirements

- Python 3.10 or later is required by the source syntax.
- Python 3.11 is the recommended project environment.
- Windows PowerShell commands are shown below; equivalent virtual-environment commands can be used on other operating systems.
- Internet access is needed for the first download of the semantic sentence-transformer model used by Data Consolidation, unless the model is already cached.

### Create and activate a virtual environment

From the project root:

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
```

Install the Python dependencies:

```powershell
python -m pip install -r requirements.txt
```

The `sentence-transformers` dependency installs a PyTorch/Transformers runtime and can make the installation substantially larger than a basic pandas/Streamlit environment.

---

## 7. Authentication configuration

The main pipeline, source environment controller, project cleaner, and Streamlit authentication use:

- `EDI_ADMIN_USERNAME`
- `EDI_ADMIN_PASSWORD_HASH`

The hash expected by the supplied source is a SHA-256 hexadecimal digest of the password's UTF-8 bytes.

For the current PowerShell session, set a username and generate the hash interactively:

```powershell
$env:EDI_ADMIN_USERNAME = "admin"

$env:EDI_ADMIN_PASSWORD_HASH = (
    python -c "import getpass, hashlib; print(hashlib.sha256(getpass.getpass('Admin password: ').encode('utf-8')).hexdigest())"
)
```

The password itself is not placed in the command text. The resulting environment variables apply to the current PowerShell session and child processes launched from it. Configure persistent Windows environment variables separately if they are needed across sessions.

**Security note:** the supplied implementation uses plain SHA-256 hashing without a password-specific salt or adaptive password-hashing function. Treat this authentication mechanism as a local-project safeguard, not production-grade identity management. Do not reuse an important account password.

`config.py` also defines `ENV_USERNAME` and `ENV_PASSWORD` settings, but the source-controller and Streamlit authentication functions specifically read `EDI_ADMIN_USERNAME` and `EDI_ADMIN_PASSWORD_HASH`.

The source code defines an `.env` path setting, but the reviewed Python files do not load `.env` files with `python-dotenv`. Setting environment variables in the shell or operating system is therefore the documented configuration method.

---

## 8. Running the application

Run commands from the project root with the virtual environment activated.

### A. Run the complete EDI pipeline

First configure authentication and select a dataset. Then run:

```powershell
python main.py
```

The program authenticates the operator, initializes required directories, resolves the input dataset, and executes the pipeline. On successful completion, it asks whether to launch the Streamlit dashboard.

### B. Run the main Streamlit dashboard

```powershell
streamlit run streamlit_app/app.py
```

The dashboard reads pipeline run manifests and allowlisted artifacts under `data/output/runs/`. It uses the EDI admin credentials configured above.

### C. Run the independent workbook intelligence application

```powershell
streamlit run streamlit_app/independent_intelligence_app.py
```

This application accepts Excel workbooks (`.xlsx`, `.xlsm`, `.xls`), profiles their worksheets, identifies candidate entity/company fields, generates adaptive charts, and produces evidence-grounded recommendations.

### D. Run Data Consolidation

Place the source company datasets in `Data_Consolidation/input/`, then run:

```powershell
python Data_Consolidation/consolidate_company_datasets.py
```

The script writes the consolidated workbook to `Data_Consolidation/output/`.

### E. Generate a Power BI project only

```powershell
python powerbi_integration.py
```

This is not a replacement for the full pipeline. It searches for a prior run with valid completed cleaning, preprocessing, analysis, forecasting, and recommendation artifacts, verifies the source dataset against the recorded hash, reuses the upstream artifacts, and executes the Power BI stage in a new run context.

A valid completed pipeline run and the original dataset file must still be available.

### F. Inspect or initialize environment paths

```powershell
python env.py --paths
python env.py --init
```

- `--paths` displays configured paths and environment information.
- `--init` creates the configured project directories.

### G. Source-code propagation and restoration

```powershell
python env.py
```

With no command-line argument, `env.py` runs the authenticated source-propagation workflow. It applies the configured `GLOBAL_CHANGES` and `TARGETED_CHANGES`, creates a source backup, validates Python syntax, compares AST structure, and can ask for approval or restore the backup.

In the supplied `env.py`, both change lists are initially empty. Nevertheless, the propagation workflow creates a backup before processing them. Do not run this command casually.

To restore the latest source backup:

```powershell
python env.py --restore
```

This operation also requires authentication and asks for confirmation.

### H. Inspect project cleanup safely

Start with a dry run:

```powershell
python project_cleaner.py --dry-run
```

To set a retention count:

```powershell
python project_cleaner.py --keep 5 --dry-run
```

After reviewing the proposed plan, run without `--dry-run` to permit the interactive deletion workflow:

```powershell
python project_cleaner.py --keep 5
```

The cleaner requires authentication and confirmation before deletion. Review the displayed plan carefully. It is intended for managed generated artifacts, not arbitrary source or input-data cleanup.

---

## 9. Important environment variables

The following variables are read directly by the supplied source code.

| Variable | Purpose |
|---|---|
| `EDI_ADMIN_USERNAME` | Username used by the source controller, main pipeline, cleaner, and Streamlit authentication. |
| `EDI_ADMIN_PASSWORD_HASH` | SHA-256 password hash expected by those authentication checks. |
| `EDI_DATASET_PATH` | Explicit dataset path; this is the variable named in `config.py` for dataset selection. |
| `EDI_PARALLEL_ANALYTICS` | Enables/disables concurrent analysis and forecasting. Defaults to enabled. Set to `0` to disable. |
| `EDI_OPEN_POWERBI` | Enables the configured Power BI Desktop launch path when set to a truthy value. |
| `EDI_AI_ENABLED` | Enables AI recommendation integration in the recommendation engine when configured as truthy. |
| `EDI_AI_API_URL` | AI provider endpoint used by the recommendation engine. |
| `EDI_AI_API_KEY` | Optional provider API key read by the recommendation engine. |
| `EDI_AI_MODEL` | Model identifier sent by the recommendation engine. |
| `EDI_AI_ENRICHMENT_ENABLED` | Controls the separate AI-enrichment setting in the recommendation engine. |
| `EDI_AI_TIMEOUT_SECONDS` | AI request timeout setting. |
| `EDI_AI_TEMPERATURE` | AI generation temperature setting. |
| `EDI_AI_MAX_OUTPUT_TOKENS` | AI output-token limit setting. |
| `EDI_AI_BATCH_SIZE` | AI recommendation batch-size setting. |
| `EDI_AI_PARALLELISM` | AI request parallelism setting. |
| `EDI_AI_MAX_RETRIES` | AI request retry setting. |
| `EDI_ANALYSIS_OUTPUT_DIRECTORY` | Overrides the analytical output directory used by `src/analysis.py`. |
| `EDI_ANALYTICAL_SOURCE` | Overrides the analytical source used by `src/analysis.py`. |
| `EDI_POWERBI_RECOVERY_MODE` | Selects the Power BI recovery output path when enabled. |

The independent Streamlit workbook application uses a local Ollama HTTP endpoint at `http://localhost:11434/api/generate` and the model name `qwen3:0.6b` in its source constants. This is a separate local service integration; it is not installed by `pip install -r requirements.txt`.

Power BI Desktop is also an external desktop application. The Python code can generate project artifacts without making Power BI Desktop a Python package dependency; opening the generated project requires a compatible Power BI Desktop installation.

---

## 10. Generated outputs and run isolation

The main orchestrator creates a new run identifier and a run-specific workspace under:

```text
data/output/runs/<RUN_ID>/
```

The stage artifact contract defined in `main.py` includes:

| Stage | Expected artifacts |
|---|---|
| Cleaning | `processed/processed_dataset.json`, `processed/processed_dataset.xlsx`, `processed/cleaning_report.json` |
| Preprocessing | `output/analytical_dataset.json`, `output/analytical_dataset.xlsx`, `output/preprocessing_report.json` |
| Analysis | `output/analysis_results.json`, `output/analysis_results.xlsx`, `output/analysis_report.json` |
| Forecasting | `output/forecast_results.json`, `output/forecast_results.xlsx`, `output/forecast_report.json` |
| Recommendation | `output/recommendation_results.xlsx`, `output/recommendation_report.json` |

The Power BI integration stages its own data and project files under the configured Power BI workspace. The source code generates PBIP/PBIR/TMDL project structures and associated metadata. Exact contents depend on the input workbook schema and the run.

Run manifests and artifact validation are used by the dashboard, cache/reuse logic, and Power BI-only workflow.

---

## 11. Dependency audit

The reviewed Python source imports these direct third-party packages:

| Package | Why it is included |
|---|---|
| `numpy` | Numerical operations in analysis, preprocessing, forecasting, and validation. |
| `pandas` | Dataset ingestion, table operations, analytics, Excel/Parquet workflows, and dashboard data handling. |
| `openpyxl` | Excel workbook reading/writing and workbook formatting. |
| `matplotlib` | Imported by the Power BI integration module. |
| `streamlit` | Dashboard and independent workbook application. |
| `plotly` | Interactive dashboard visualizations. |
| `sentence-transformers` | Semantic matching of source column headers in Data Consolidation. |
| `huggingface-hub` | Model-cache lookup and download for semantic matching. |

Two additional packages are included for pandas' file-format engines:

- `xlrd` for legacy `.xls` workbook reading.
- `pyarrow` for Parquet reading.

The sentence-transformers stack installs its own runtime dependencies, including PyTorch/Transformers-related packages.

The previous requirements file also listed `seaborn`, `scikit-learn`, `scipy`, `joblib`, `python-dotenv`, `xlsxwriter`, and `fpdf2`. The reviewed application source does not directly import these packages. Some may arrive transitively through other packages, but they are not listed as direct project requirements here.

The requirements file uses version ranges rather than exact lockfile pins. For a fully reproducible deployment, generate and maintain a tested lock file from the actual target Python environment after installation and end-to-end validation.

---

## 12. External services and optional integrations

### Hugging Face model download

The Data Consolidation module uses the model `sentence-transformers/all-MiniLM-L6-v2` for semantic column-header matching. It checks the local Hugging Face cache first and attempts an online download if the model is not cached. The source sets `SEMANTIC_MODEL_REQUIRED = True` in its active configuration, so missing model dependencies or an unavailable model can stop consolidation.

### Ollama

The independent Streamlit intelligence application calls a local Ollama HTTP endpoint and specifies `qwen3:0.6b`. Ollama itself is not a Python package in `requirements.txt`. Install and run Ollama separately and ensure the configured local service/model is available if this feature is to be used.

### AI recommendation endpoint

`src/recommendation_engine.py` can call an externally configured HTTP AI endpoint using the `EDI_AI_*` settings. The source also contains deterministic recommendation fallback logic. The endpoint/provider is not fixed by the Python dependency file.

### Power BI Desktop

Power BI Desktop is an optional external application used to open generated Power BI projects. Its installation is separate from the Python environment.

---

## 13. Operational and security notes

1. Run the project from its root directory so local imports and relative project paths resolve consistently.
2. Use one selected main-pipeline dataset at a time. If automatic discovery is used, keep exactly one supported dataset file in `data/raw/`.
3. The Data Consolidation utility and the main pipeline are separate workflows. Consolidation output must be selected explicitly for the main pipeline.
4. The main pipeline, source controller, cleaner, and Streamlit dashboard use the EDI admin environment variables. Configure them before starting those components.
5. The first semantic-model download requires network access unless the model is already cached.
6. The Ollama and Power BI Desktop integrations require their external applications/services separately.
7. Review cleanup plans before approving deletion. Keep backups of important source data and generated outputs.
8. `env.py` performs literal source-text replacement, not semantic refactoring. Its AST comparison is a risk signal, not proof that runtime behaviour is unchanged.
9. The current `env.py` global discovery and targeted-modification paths do not apply exactly the same eligibility checks. Treat source propagation as a sensitive maintenance operation until that policy is unified.
10. The current source authentication uses unsalted SHA-256. Do not treat it as production-grade authentication.
11. The repository source was statically inspected for this documentation. Syntax parsing is not a substitute for running every workflow against representative datasets and external services.

---

## 14. Suggested validation sequence after installation

Use a copy of representative input data and a development environment.

1. Confirm imports and package installation.
2. Run `python env.py --paths`.
3. Confirm the selected dataset path and supported format.
4. Run the Data Consolidation utility separately if a combined workbook is required.
5. Run the full pipeline with `python main.py`.
6. Review the run manifest and every stage's expected artifacts.
7. Launch `streamlit run streamlit_app/app.py` and confirm dashboard access and run selection.
8. Test the independent workbook app separately if it is part of the intended workflow.
9. Run `python powerbi_integration.py` only after a valid completed pipeline run exists.
10. Test `project_cleaner.py --dry-run` before approving any cleanup.

---

## 15. Current review scope

This documentation was prepared from the supplied source archive containing 23 Python files, including `env.py`, and the project tree shared during the conversation.

All 23 Python files were parsed for Python syntax during the static review. Imports, module responsibilities, file-format readers, environment-variable references, CLI entry points, and the main pipeline artifact contract were inspected.

No full end-to-end execution is claimed here. Such a test would require the target Python environment, credentials, a representative dataset, and (for the relevant features) the semantic model cache, Ollama service, or Power BI Desktop.
