"""
Global project configuration.

Dataset structure and business meaning are NOT defined here.
They are discovered at runtime from the supplied dataset.

Only technical settings are fixed:
- supported input file formats
- maximum input file size
- runtime/resource safeguards
- project/output locations
"""

from __future__ import annotations

import os
from pathlib import Path


# ============================================================
# PROJECT
# ============================================================

PROJECT_NAME = "Enterprise Demand Intelligence Platform"
PROJECT_VERSION = "1.0.0"
PROJECT_ROOT = Path(__file__).resolve().parent


# ============================================================
# DIRECTORIES
# ============================================================

DATA_DIRECTORY = PROJECT_ROOT / "data"
RAW_DATA_DIRECTORY = DATA_DIRECTORY / "raw"
PROCESSED_DATA_DIRECTORY = DATA_DIRECTORY / "processed"
OUTPUT_DATA_DIRECTORY = DATA_DIRECTORY / "output"
UPLOAD_DIRECTORY = DATA_DIRECTORY / "uploads"

DOCUMENTATION_DIRECTORY = PROJECT_ROOT / "documentation"
POWERBI_DIRECTORY = PROJECT_ROOT / "powerbi"
POWERBI_DATA_DIRECTORY = POWERBI_DIRECTORY / "data"
POWERBI_DOCUMENTATION_DIRECTORY = POWERBI_DIRECTORY / "documentation"

SOURCE_DIRECTORY = PROJECT_ROOT / "src"
LOG_DIRECTORY = PROJECT_ROOT / "logs"
ARCHIVE_DIRECTORY = PROJECT_ROOT / "_archive"
BACKUP_DIRECTORY = PROJECT_ROOT / "backups"


# ============================================================
# DATASET INPUT
# ============================================================
#
# IMPORTANT:
# No dataset filename is hardcoded here.
#
# The loader resolves the active file in this order:
#   1. explicit command-line path
#   2. ACTIVE_DATASET_PATH, if configured
#   3. EDI_DATASET_PATH environment variable
#   4. automatic discovery of one supported file in data/raw
#
# Therefore VS Code's "Run Python File" button works without
# requiring a command-line argument.
# ============================================================

ACTIVE_DATASET_PATH = None

DATASET_PATH_ENV_VAR = "EDI_DATASET_PATH"

SUPPORTED_DATASET_FORMATS = {
    "json",
    "csv",
    "xlsx",
    "xls",
    "parquet",
}

MAX_DATASET_FILE_SIZE_MB = 500


# ============================================================
# DATASET DISCOVERY
# ============================================================

DATASET_DISCOVERY_ENABLED = True
AUTOMATIC_SCHEMA_DISCOVERY = True
AUTOMATIC_TABLE_DISCOVERY = True
AUTOMATIC_COLUMN_DISCOVERY = True
AUTOMATIC_TYPE_DISCOVERY = True
AUTOMATIC_RELATIONSHIP_DISCOVERY = True


# ============================================================
# TECHNICAL RESOURCE SAFEGUARDS
# ============================================================
#
# These limits protect the application from excessive resource
# consumption. They do not describe what a valid business
# dataset must contain.
# ============================================================

MAX_DISCOVERED_TABLES = 1000
MAX_DISCOVERED_COLUMNS_PER_TABLE = 1000
MAX_TOTAL_ROWS = 500_000_000
SAMPLE_SIZE = 5000
RELATIONSHIP_SAMPLE_SIZE = 2000
MAX_RELATIONSHIP_PAIRS = 5000

# Demand-sensing relationship analysis limits.
# Technical/model-computation control only; no dataset-specific assumptions.
DEMAND_MIN_OBSERVATIONS = 30
DEMAND_MAX_CANDIDATES_PER_TABLE = 10
DEMAND_MAX_SIGNALS_PER_CANDIDATE = 20
DEMAND_MAX_LAG_PERIODS = 7
DEMAND_MIN_ABS_CORRELATION = 0.20
DEMAND_MIN_DIRECTIONAL_AGREEMENT = 0.55
DEMAND_MAX_RELATIONSHIP_COMPARISONS = 5000

# Technical forecasting resource limits.
FORECAST_MAX_POINTS = 1000
FORECAST_MAX_CANDIDATES = 500
FORECAST_MAX_GENERATED_SERIES = 50000


# ============================================================
# OPTIONAL COMPATIBILITY SETTINGS
# ============================================================
#
# These names are retained only so older project modules that
# import them do not immediately fail.
#
# They contain no dataset-specific names or values.
# ============================================================

PROCESSING_MODE = "dataset_driven"
DATASET_SCHEMA_MODE = "dynamic"
DATASET_COMPATIBILITY_MODE = "capability_based"

SUPPORTED_PROCESSING_MODES = {"dataset_driven"}
SUPPORTED_DATASET_SCHEMA_MODES = {
    "dynamic",
    "reference_compatible",
}
SUPPORTED_DATASET_COMPATIBILITY_MODES = {
    "strict_reference",
    "reference_compatible",
    "capability_based",
}

PRIMARY_KEY_MAP = {}
FOREIGN_KEY_MAP = {}
REFERENCE_PRIMARY_KEY_MAP = {}
REFERENCE_FOREIGN_KEY_MAP = {}
REFERENCE_EXPECTED_COLUMNS = {}

EXPECTED_TABLE_COUNT = None
REFERENCE_DATASET_TABLE_COUNT = None
REFERENCE_DATASET_TABLES = set()

PERCENTAGE_COLUMNS = set()
FLAG_COLUMNS = set()
NON_NEGATIVE_COLUMNS = set()

DEMAND_TARGET_COLUMN = None
DEMAND_SIGNAL_COLUMN = None
DEMAND_CHANGE_COLUMN = None


# ============================================================
# APPLICATION / PIPELINE SETTINGS
# ============================================================

ETL_ENABLED = True
ETL_ENCODING = "utf-8"
ETL_READ_JSON_ENCODING = "utf-8"
ETL_CREATE_RAW_WORKBOOK = True
ETL_CREATE_CLEANED_WORKBOOK = True
ETL_OVERWRITE_OUTPUTS = True
ETL_AUTOMATIC_SCHEMA_DISCOVERY = True
ETL_AUTOMATIC_COLUMN_MAPPING = True
ETL_AUTOMATIC_TABLE_DISCOVERY = True
ETL_PRESERVE_UNKNOWN_TABLES = True
ETL_PRESERVE_UNKNOWN_COLUMNS = True

STANDARDIZATION_ENABLED = True
COLUMN_MAPPING_ENABLED = True
TABLE_MAPPING_ENABLED = True
TYPE_STANDARDIZATION_ENABLED = True
DATE_STANDARDIZATION_ENABLED = True
UNIT_STANDARDIZATION_ENABLED = True

ANALYSIS_ENABLED = True
ANALYSIS_MODE = "dataset_driven"
ANALYSIS_SUPPORTS_CUSTOM_DATASET = True

DEMAND_SENSING_ENABLED = True
FORECASTING_ENABLED = True
FORECASTING_MINIMUM_HISTORY_DAYS = 30

RECOMMENDATION_ENABLED = True
RECOMMENDATION_MODE = "configurable"

AI_ENABLED = True
AI_PROVIDER = "configurable"
AI_MODEL_NAME = "configurable"
AI_TEMPERATURE = 0.2
AI_MAX_OUTPUT_TOKENS = 2000


# ============================================================
# OUTPUT LOCATIONS
# ============================================================

RAW_WORKBOOK_PATH = RAW_DATA_DIRECTORY / "raw_dataset.xlsx"

PROCESSED_WORKBOOK_PATH = (
    PROCESSED_DATA_DIRECTORY / "processed_dataset.xlsx"
)

CLEANED_WORKBOOK_PATH = PROCESSED_WORKBOOK_PATH
PROCESSED_WORKBOOK = PROCESSED_WORKBOOK_PATH

ANALYTICAL_WORKBOOK_PATH = (
    OUTPUT_DATA_DIRECTORY / "analytical_dataset.xlsx"
)

ANALYTICAL_WORKBOOK = ANALYTICAL_WORKBOOK_PATH
ANALYTICAL_OUTPUT_DIRECTORY = OUTPUT_DATA_DIRECTORY

POWERBI_PRIMARY_DATA_SOURCE = PROCESSED_WORKBOOK_PATH


# ============================================================
# UPLOAD / DEPLOYMENT
# ============================================================

DATA_UPLOAD_ENABLED = True
ALLOW_USER_DATASET_REPLACEMENT = True
KEEP_PREVIOUS_DATASET_BACKUP = True

APPLICATION_HOST = "127.0.0.1"
APPLICATION_PORT = 8000
APPLICATION_DEBUG = True
APPLICATION_NAME = PROJECT_NAME

DEPLOYMENT_ENVIRONMENT = "local"
SUPPORTED_DEPLOYMENT_ENVIRONMENTS = {
    "local",
    "development",
    "production",
    "cloud",
}


# ============================================================
# VALIDATION BEHAVIOUR
# ============================================================

REQUIRE_NON_NULL_VALUES = True
REQUIRE_NON_BLANK_VALUES = True
REQUIRE_CONSISTENT_COLUMNS = True
REQUIRE_UNIQUE_PRIMARY_KEYS = False
REQUIRE_VALID_FOREIGN_KEYS = False
REQUIRE_VALID_DATES = False
REQUIRE_FINITE_NUMERIC_VALUES = True

ALLOW_ADDITIONAL_TABLES = True
ALLOW_ADDITIONAL_COLUMNS = True
ALLOW_TABLE_NAME_DIFFERENCES = True
ALLOW_COLUMN_NAME_DIFFERENCES = True
REQUIRE_ALL_REFERENCE_TABLES = False
REQUIRE_ALL_REFERENCE_COLUMNS = False


# ============================================================
# SECURITY / ENVIRONMENT
# ============================================================

ENVIRONMENT_FILE = PROJECT_ROOT / ".env"
USE_ENVIRONMENT_VARIABLES = True

ENV_USERNAME = os.getenv("ENV_USERNAME", "admin")
ENV_PASSWORD = os.getenv("ENV_PASSWORD", "")


# ============================================================
# LOGGING
# ============================================================

LOGGING_ENABLED = True
LOG_LEVEL = "INFO"
LOG_FILE = LOG_DIRECTORY / "application.log"

ENCODING = "utf-8"

SOURCE_DIRECTORIES = [SOURCE_DIRECTORY]
PYTHON_SOURCE_EXTENSIONS = {".py"}

GLOBAL_CHANGES = []
TARGETED_CHANGES = []


# ============================================================
# DIRECTORY INITIALIZATION
# ============================================================

for directory in (
    DATA_DIRECTORY,
    RAW_DATA_DIRECTORY,
    PROCESSED_DATA_DIRECTORY,
    OUTPUT_DATA_DIRECTORY,
    UPLOAD_DIRECTORY,
    DOCUMENTATION_DIRECTORY,
    POWERBI_DIRECTORY,
    POWERBI_DATA_DIRECTORY,
    POWERBI_DOCUMENTATION_DIRECTORY,
    SOURCE_DIRECTORY,
    LOG_DIRECTORY,
    ARCHIVE_DIRECTORY,
    BACKUP_DIRECTORY,
):
    directory.mkdir(parents=True, exist_ok=True)


# ============================================================
# CONFIGURATION TEST
# ============================================================

if __name__ == "__main__":
    print("=" * 70)
    print("PROJECT CONFIGURATION")
    print("=" * 70)
    print(f"Project root       : {PROJECT_ROOT}")
    print(f"Raw data directory : {RAW_DATA_DIRECTORY}")
    print(
        "Active dataset     : "
        f"{ACTIVE_DATASET_PATH or 'AUTO-DISCOVER'}"
    )
    print(
        "Supported formats  : "
        f"{', '.join(sorted(SUPPORTED_DATASET_FORMATS))}"
    )
    print(
        f"Maximum file size  : {MAX_DATASET_FILE_SIZE_MB} MB"
    )
    print("=" * 70)
