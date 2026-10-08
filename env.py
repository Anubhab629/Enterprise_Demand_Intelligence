"""
Enterprise Demand Intelligence Platform
SOURCE ENVIRONMENT CONTROLLER + PROJECT ENVIRONMENT CONFIGURATION

Purpose
-------
This module provides two independent capabilities:

1. Centralized, dataset-independent project environment/path configuration.
2. Controlled physical modification of Python source code.

Architecture
------------
The application is DATASET-DRIVEN.

This controller contains no assumptions about:
    - the current demo dataset
    - company count
    - table names
    - column names
    - business entities
    - business-specific transformations

Dataset information is discovered by the application at runtime.

The source controller modifies Python source files only. Dataset files
are protected from source-code modification operations.

IMPORTANT
---------
Edit only:
    CURRENT_VALUES
    GLOBAL_CHANGES
    TARGETED_CHANGES

GLOBAL_CHANGES:
    Apply an exact textual replacement across every eligible Python
    source file.

TARGETED_CHANGES:
    Apply an exact textual replacement to one eligible Python file,
    optionally restricted by occurrence or line number.

Every propagation operation:
    1. authenticates the operator
    2. discovers eligible Python source files
    3. creates a backup BEFORE modification
    4. applies physical source changes
    5. validates Python syntax
    6. compares before/after AST structure
    7. classifies possible behavior risk
    8. asks for confirmation when appropriate
    9. restores the backup when the operator rejects risky changes
       or syntax validation fails

The changes are real filesystem changes. They are NOT merely displayed
in the terminal.
"""

from __future__ import annotations

import ast
import difflib
import getpass
import hashlib
import hmac
import os
import shutil
import sys
from collections import Counter
from pathlib import Path
from typing import Any


# ============================================================
# PROJECT LOCATION
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parent


# ============================================================
# GLOBAL VALUES
# ============================================================

CURRENT_VALUES: dict[str, Any] = {
    "PROJECT_NAME": "Enterprise Demand Intelligence System",
    "PROJECT_VERSION": "1.0.0",

    # Authentication settings are supplied through environment variables.`r`n    # No deployment credential is stored in source control.`r`n
    # Generic project directories.
    "DATA_DIRECTORY": "data",
    "RAW_DATA_DIRECTORY": "data/raw",
    "PROCESSED_DATA_DIRECTORY": "data/processed",
    "REFERENCE_DATA_DIRECTORY": "data/reference",
    "INPUT_DATA_DIRECTORY": "data/input",
    "COMPANY_DATA_DIRECTORY": "data/companies",
    "DOCUMENTATION_DIRECTORY": "documentation",
    "OUTPUTS_DIRECTORY": "outputs",
    "POWERBI_DIRECTORY": "powerbi",

    # Resource-protection limits.
    # These are limits, NOT dataset schema assumptions.
    "MAX_ACCESS_BOUNDARY_CARDINALITY": 100,
    "MAX_DATASET_FILE_SIZE_MB": 500,
    "MAX_COMPANIES_PER_DATASET": 25,

    # Default/current input settings.
    #
    # These identify the currently active input instance only.
    # They do not define the application schema and are not used
    # by source-modification logic.
    "RAW_JSON_NAME": "dataset.json",
    "RAW_WORKBOOK_NAME": "raw_dataset.xlsx",
    "PROCESSED_WORKBOOK_NAME": "processed_dataset.xlsx",
    "ANALYTICAL_FEATURES_WORKBOOK_NAME":
        "analytical_features.xlsx",
    "RAW_ARCHIVE_NAME": "dataset_archive.zip",

    "ACTIVE_DATASET_FORMAT": "json",
    "ACTIVE_DATASET_PATH": "data/raw/dataset.json",

    # Generic application labels/messages.
    "PRIMARY_KEY_LABEL": "Primary Key",
    "FOREIGN_KEY_LABEL": "Foreign Key",
    "ETL_LABEL": "ETL Pipeline",

    "APPLICATION_INITIALIZED_MESSAGE":
        "Application initialized successfully.",
    "ACCESS_GRANTED_MESSAGE":
        "Access granted. Continuing application processing...",
    "ACCESS_DENIED_MESSAGE":
        "Access denied. Application terminated.",

    "ETL_START_MESSAGE": "Starting ETL pipeline...",
    "SECTION_SEPARATOR_LENGTH": "60",
    "ETL_LOAD_MESSAGE": "1. Loading raw dataset...",
    "ETL_STANDARDIZE_MESSAGE": "2. Standardizing data types...",
    "ETL_CLEAN_MESSAGE": "3. Cleaning dataset...",
    "ETL_PRIMARY_KEY_MESSAGE": "4. Validating primary keys...",
    "ETL_FOREIGN_KEY_MESSAGE": "5. Validating foreign keys...",
    "ETL_SAVE_MESSAGE": "6. Saving processed dataset...",
    "ETL_COMPLETED_MESSAGE":
        "ETL pipeline completed successfully.",
    "PROCESSED_WORKBOOK_MESSAGE": "Processed workbook:",

    "DATASET_UPLOAD_MESSAGE":
        "Customer dataset uploaded successfully.",
    "DATASET_VALIDATION_MESSAGE":
        "Validating uploaded dataset...",
    "DATASET_COMPANY_COUNT_MESSAGE":
        "Companies detected in dataset:",
    "DATASET_LIMIT_EXCEEDED_MESSAGE":
        "Dataset exceeds the configured complexity limit.",
    "DATASET_ACTIVATED_MESSAGE":
        "Customer dataset activated successfully.",
}


# ============================================================
# SOURCE CHANGE CONFIGURATION
# ============================================================

# Exact global replacements across eligible Python source files.
GLOBAL_CHANGES: list[dict[str, Any]] = []

# Exact replacements in one eligible Python source file.
TARGETED_CHANGES: list[dict[str, Any]] = []


# ============================================================
# ENVIRONMENT ACCESS
# ============================================================

ENV = CURRENT_VALUES


# ============================================================
# CENTRALIZED PATH CONFIGURATION
# ============================================================

DATA_DIRECTORY = PROJECT_ROOT / CURRENT_VALUES["DATA_DIRECTORY"]
RAW_DATA_DIRECTORY = PROJECT_ROOT / CURRENT_VALUES["RAW_DATA_DIRECTORY"]
PROCESSED_DATA_DIRECTORY = (
    PROJECT_ROOT / CURRENT_VALUES["PROCESSED_DATA_DIRECTORY"]
)
REFERENCE_DATA_DIRECTORY = (
    PROJECT_ROOT / CURRENT_VALUES["REFERENCE_DATA_DIRECTORY"]
)
INPUT_DATA_DIRECTORY = PROJECT_ROOT / CURRENT_VALUES["INPUT_DATA_DIRECTORY"]
COMPANY_DATA_DIRECTORY = (
    PROJECT_ROOT / CURRENT_VALUES["COMPANY_DATA_DIRECTORY"]
)
DOCUMENTATION_DIRECTORY = (
    PROJECT_ROOT / CURRENT_VALUES["DOCUMENTATION_DIRECTORY"]
)
OUTPUTS_DIRECTORY = PROJECT_ROOT / CURRENT_VALUES["OUTPUTS_DIRECTORY"]
POWERBI_DIRECTORY = PROJECT_ROOT / CURRENT_VALUES["POWERBI_DIRECTORY"]


# ============================================================
# DATASET PATH CONFIGURATION
# ============================================================

RAW_JSON_PATH = RAW_DATA_DIRECTORY / CURRENT_VALUES["RAW_JSON_NAME"]
RAW_WORKBOOK_PATH = RAW_DATA_DIRECTORY / CURRENT_VALUES["RAW_WORKBOOK_NAME"]
PROCESSED_WORKBOOK_PATH = (
    PROCESSED_DATA_DIRECTORY
    / CURRENT_VALUES["PROCESSED_WORKBOOK_NAME"]
)
ANALYTICAL_FEATURES_WORKBOOK_PATH = (
    PROCESSED_DATA_DIRECTORY
    / CURRENT_VALUES["ANALYTICAL_FEATURES_WORKBOOK_NAME"]
)
RAW_ARCHIVE_PATH = RAW_DATA_DIRECTORY / CURRENT_VALUES["RAW_ARCHIVE_NAME"]

ACTIVE_DATASET_PATH = (
    PROJECT_ROOT / CURRENT_VALUES["ACTIVE_DATASET_PATH"]
)

# ============================================================
# DATASET FORMAT CONFIGURATION
# ============================================================

SUPPORTED_DATASET_FORMATS = {
    "json",
    "csv",
    "xlsx",
    "xls",
    "parquet",
}

DATASET_EXTENSIONS = {
    ".csv",
    ".json",
    ".xlsx",
    ".xls",
    ".parquet",
}


def get_active_dataset_path() -> Path:
    """Return the absolute path of the currently active dataset."""
    return ACTIVE_DATASET_PATH


def get_active_dataset_format() -> str:
    """Return the configured active dataset format."""
    return str(CURRENT_VALUES["ACTIVE_DATASET_FORMAT"])


def get_max_access_boundary_cardinality() -> int:
    """Return the generic technical cardinality limit for access-boundary discovery."""
    return int(CURRENT_VALUES["MAX_ACCESS_BOUNDARY_CARDINALITY"])



def get_max_dataset_file_size_mb() -> int:
    """Return the configured maximum dataset file size in MB."""
    return int(CURRENT_VALUES["MAX_DATASET_FILE_SIZE_MB"])


def is_supported_dataset_format(file_path: Path) -> bool:
    """Return True when the file extension is supported."""
    return file_path.suffix.lower() in DATASET_EXTENSIONS


def validate_dataset_file_size(file_path: Path) -> bool:
    """Validate dataset size against the generic resource limit."""
    if not file_path.exists() or not file_path.is_file():
        return False

    maximum_bytes = (
        get_max_dataset_file_size_mb()
        * 1024
        * 1024
    )

    return file_path.stat().st_size <= maximum_bytes


# ============================================================
# PATH HELPERS
# ============================================================

def get_project_root() -> Path:
    """Return the absolute project root."""
    return PROJECT_ROOT


def get_raw_data_directory() -> Path:
    """Return the raw data directory."""
    return RAW_DATA_DIRECTORY


def get_processed_data_directory() -> Path:
    """Return the processed data directory."""
    return PROCESSED_DATA_DIRECTORY


def get_input_data_directory() -> Path:
    """Return the external dataset input/upload directory."""
    return INPUT_DATA_DIRECTORY


def get_company_data_directory() -> Path:
    """Return the customer/company data directory."""
    return COMPANY_DATA_DIRECTORY


def get_powerbi_directory() -> Path:
    """Return the Power BI directory."""
    return POWERBI_DIRECTORY


# ============================================================
# DIRECTORY INITIALIZATION
# ============================================================

def ensure_project_directories() -> None:
    """
    Create configured project directories when absent.

    This function does not alter dataset contents.
    """
    directories = (
        DATA_DIRECTORY,
        RAW_DATA_DIRECTORY,
        PROCESSED_DATA_DIRECTORY,
        REFERENCE_DATA_DIRECTORY,
        INPUT_DATA_DIRECTORY,
        COMPANY_DATA_DIRECTORY,
        DOCUMENTATION_DIRECTORY,
        OUTPUTS_DIRECTORY,
        POWERBI_DIRECTORY,
    )

    for directory in directories:
        directory.mkdir(parents=True, exist_ok=True)


# ============================================================
# DATASET PROTECTION
# ============================================================

def is_dataset_file(path: Path) -> bool:
    """Return True when a path has a protected dataset extension."""
    return path.suffix.lower() in DATASET_EXTENSIONS


# ============================================================
# SOURCE-CODE PROTECTION
# ============================================================

BACKUP_DIRECTORY = PROJECT_ROOT / ".env_backups"

EXCLUDED_FILES = {
    "env.py",
    "env_manager.py",
    "source_modifier.py",
}

EXCLUDED_DIRECTORIES = {
    "__pycache__",
    ".git",
    ".venv",
    "venv",
    "env",
    ".env_backups",
    "migration_backup",
}


def is_excluded_source_path(path: Path) -> bool:
    """
    Return True when a source path is protected from modification.

    Both the filename and every path component are checked.
    """
    if path.name in EXCLUDED_FILES:
        return True

    return any(
        part in EXCLUDED_DIRECTORIES
        for part in path.parts
    )


def is_eligible_python_source(path: Path) -> bool:
    """Return True only for eligible Python source files."""
    return (
        path.suffix.lower() == ".py"
        and not is_dataset_file(path)
        and not is_excluded_source_path(path)
        and path.is_file()
    )


# ============================================================
# AUTHENTICATION
# ============================================================

def hash_password(password: str) -> str:
    """Return the SHA-256 hash of a supplied password."""
    return hashlib.sha256(
        password.encode("utf-8")
    ).hexdigest()


def authenticate() -> bool:
    """Authenticate the operator before source modification."""
    print()
    print("=" * 70)
    print(CURRENT_VALUES["PROJECT_NAME"])
    print("SOURCE ENVIRONMENT CONTROLLER")
    print("=" * 70)

    username = input("Username: ").strip()
    password = getpass.getpass("Password: ")

    expected_username = os.getenv("EDI_ADMIN_USERNAME", "").strip()
    expected_hash = os.getenv("EDI_ADMIN_PASSWORD_HASH", "").strip()

    if not expected_username or not expected_hash:
        print()
        print(
            "Authentication configuration is missing. "
            "Set EDI_ADMIN_USERNAME and EDI_ADMIN_PASSWORD_HASH."
        )
        return False

    supplied_hash = hash_password(password)

    if (
        username == expected_username
        and hmac.compare_digest(
            supplied_hash,
            expected_hash,
        )
    ):
        print()
        print("Authentication successful.")
        return True

    print()
    print("Authentication failed.")
    return False


# ============================================================
# SOURCE DISCOVERY
# ============================================================

def get_source_files() -> list[Path]:
    """
    Return every eligible Python source file in the project.

    Dataset files are never returned.
    """
    source_files: list[Path] = []

    for path in PROJECT_ROOT.rglob("*.py"):
        if is_eligible_python_source(path):
            source_files.append(path)

    return sorted(source_files)


# ============================================================
# BACKUP
# ============================================================

def create_backup(source_files: list[Path]) -> Path:
    """
    Create a complete backup of eligible Python source files.

    The backup is created before any source modification.
    """
    BACKUP_DIRECTORY.mkdir(parents=True, exist_ok=True)

    backup_dir = BACKUP_DIRECTORY / "backup"
    counter = 1

    while backup_dir.exists():
        backup_dir = BACKUP_DIRECTORY / f"backup_{counter}"
        counter += 1

    backup_dir.mkdir(parents=True, exist_ok=False)

    for source_file in source_files:
        relative_path = source_file.relative_to(PROJECT_ROOT)
        destination = backup_dir / relative_path

        destination.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        shutil.copy2(source_file, destination)

    return backup_dir


# ============================================================
# BACKUP RESTORATION
# ============================================================

def restore_backup(backup_dir: Path) -> None:
    """Restore Python source files from a backup directory."""
    if not backup_dir.exists():
        raise FileNotFoundError(
            f"Backup directory does not exist: {backup_dir}"
        )

    for backup_file in backup_dir.rglob("*.py"):
        relative_path = backup_file.relative_to(backup_dir)
        destination = PROJECT_ROOT / relative_path

        destination.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        shutil.copy2(backup_file, destination)


# ============================================================
# PYTHON VALIDATION
# ============================================================

def validate_python_sources() -> tuple[bool, list[str]]:
    """Validate Python syntax for every eligible source file."""
    errors: list[str] = []

    for source_file in get_source_files():
        try:
            source = source_file.read_text(
                encoding="utf-8-sig"
            )

            ast.parse(
                source,
                filename=str(source_file),
            )

        except Exception as exc:
            errors.append(
                f"{source_file.relative_to(PROJECT_ROOT)}: {exc}"
            )

    return len(errors) == 0, errors


# ============================================================
# SOURCE READING / WRITING HELPERS
# ============================================================

def read_source_file(path: Path) -> str:
    """Read an eligible Python source file."""
    return path.read_text(encoding="utf-8")


def write_source_file(path: Path, source: str) -> None:
    """Write modified Python source to the physical filesystem."""
    path.write_text(source, encoding="utf-8")


# ============================================================
# GLOBAL REPLACEMENT
# ============================================================

def replace_global(
    old_value: str,
    new_value: str,
) -> list[Path]:
    """
    Physically replace an exact value across eligible Python files.

    Returns the files that were actually modified.
    """
    if not old_value or old_value == new_value:
        return []

    changed_files: list[Path] = []

    for source_file in get_source_files():
        try:
            original = read_source_file(source_file)
        except UnicodeDecodeError:
            continue

        if old_value not in original:
            continue

        updated = original.replace(
            old_value,
            new_value,
        )

        if updated == original:
            continue

        write_source_file(
            source_file,
            updated,
        )

        changed_files.append(source_file)

    return changed_files


# ============================================================
# TARGETED REPLACEMENT
# ============================================================

def replace_targeted(
    file_name: str,
    old_value: str,
    new_value: str,
    occurrence: int | None = None,
    line_number: int | None = None,
) -> tuple[bool, str]:
    """
    Physically replace an exact value in one Python source file.

    occurrence:
        None -> replace every occurrence in that file.
        1+  -> replace only that occurrence.

    line_number:
        1+  -> restrict replacement to that exact line.
        None -> no line restriction.
    """
    target = PROJECT_ROOT / file_name

    if not target.exists():
        return False, f"File not found: {file_name}"

    if not target.is_file():
        return False, f"Target is not a file: {file_name}"

    if is_dataset_file(target):
        return (
            False,
            "Modification blocked: dataset files cannot be "
            "modified by the source controller.",
        )

    if target.suffix.lower() != ".py":
        return (
            False,
            "Targeted modification is restricted to Python files.",
        )

    if is_excluded_source_path(target):
        return (
            False,
            f"Modification blocked for protected file: {file_name}",
        )

    if not old_value:
        return False, "Old value cannot be empty."

    try:
        original = read_source_file(target)
    except UnicodeDecodeError:
        return (
            False,
            f"Unable to read UTF-8 source file: {file_name}",
        )

    # --------------------------------------------------------
    # LINE-SPECIFIC MODE
    # --------------------------------------------------------

    if line_number is not None:
        if line_number < 1:
            return False, "line_number must be 1 or greater."

        lines = original.splitlines(keepends=True)

        if line_number > len(lines):
            return (
                False,
                f"Line {line_number} does not exist in {file_name}. "
                f"File contains {len(lines)} lines.",
            )

        index = line_number - 1
        selected_line = lines[index]

        if old_value not in selected_line:
            return (
                False,
                f"Safety check failed: line {line_number} does not "
                f"contain the expected old value.",
            )

        lines[index] = selected_line.replace(
            old_value,
            new_value,
        )

        updated = "".join(lines)

        if updated == original:
            return False, "No change required."

        write_source_file(target, updated)

        return True, f"{file_name} line {line_number}"

    # --------------------------------------------------------
    # OCCURRENCE MODE
    # --------------------------------------------------------

    if old_value not in original:
        return (
            False,
            f"Term not found in {file_name}: {old_value}",
        )

    if occurrence is None:
        updated = original.replace(
            old_value,
            new_value,
        )
    else:
        if occurrence < 1:
            return False, "Occurrence must be 1 or greater."

        positions: list[int] = []
        start = 0

        while True:
            position = original.find(
                old_value,
                start,
            )

            if position == -1:
                break

            positions.append(position)
            start = position + len(old_value)

        if occurrence > len(positions):
            return (
                False,
                f"Occurrence {occurrence} does not exist in "
                f"{file_name}. Available occurrences: {len(positions)}",
            )

        position = positions[occurrence - 1]

        updated = (
            original[:position]
            + new_value
            + original[position + len(old_value):]
        )

    if updated == original:
        return False, "No change required."

    write_source_file(target, updated)

    return True, file_name


# ============================================================
# AST HELPERS
# ============================================================

def parse_python_source(
    source: str,
    filename: str,
) -> ast.AST:
    """Parse Python source into an AST."""
    return ast.parse(
        source,
        filename=filename,
    )


def normalize_ast(tree: ast.AST) -> str:
    """Return an AST representation without location metadata."""
    return ast.dump(
        tree,
        annotate_fields=True,
        include_attributes=False,
    )


# ============================================================
# IDENTIFIER ANALYSIS
# ============================================================

def get_identifier_changes(
    before_tree: ast.AST,
    after_tree: ast.AST,
) -> list[dict[str, Any]]:
    """Detect function, class, variable, argument and import changes."""
    changes: list[dict[str, Any]] = []

    def collect_identifiers(tree: ast.AST) -> dict[str, list[str]]:
        identifiers = {
            "functions": [],
            "classes": [],
            "variables": [],
            "arguments": [],
            "imports": [],
        }

        for node in ast.walk(tree):
            if isinstance(
                node,
                (ast.FunctionDef, ast.AsyncFunctionDef),
            ):
                identifiers["functions"].append(node.name)

                for argument in node.args.args:
                    identifiers["arguments"].append(argument.arg)

                for argument in node.args.posonlyargs:
                    identifiers["arguments"].append(argument.arg)

                for argument in node.args.kwonlyargs:
                    identifiers["arguments"].append(argument.arg)

                if node.args.vararg is not None:
                    identifiers["arguments"].append(
                        node.args.vararg.arg
                    )

                if node.args.kwarg is not None:
                    identifiers["arguments"].append(
                        node.args.kwarg.arg
                    )

            elif isinstance(node, ast.ClassDef):
                identifiers["classes"].append(node.name)

            elif isinstance(node, ast.Name):
                identifiers["variables"].append(node.id)

            elif isinstance(node, ast.Import):
                for alias in node.names:
                    identifiers["imports"].append(
                        alias.asname or alias.name
                    )

            elif isinstance(node, ast.ImportFrom):
                for alias in node.names:
                    identifiers["imports"].append(
                        alias.asname or alias.name
                    )

        return identifiers

    before = collect_identifiers(before_tree)
    after = collect_identifiers(after_tree)

    def add_rename_or_add_remove(
        category: str,
        before_items: list[str],
        after_items: list[str],
    ) -> None:
        before_counter = Counter(before_items)
        after_counter = Counter(after_items)

        removed = list(before_counter - after_counter)
        added = list(after_counter - before_counter)

        rename_count = min(
            len(removed),
            len(added),
        )

        for index in range(rename_count):
            old_value = removed[index]
            new_value = added[index]

            changes.append(
                {
                    "category": category,
                    "type": category,
                    "old": old_value,
                    "new": new_value,
                    "value": f"{old_value} -> {new_value}",
                }
            )

        for identifier in removed[rename_count:]:
            changes.append(
                {
                    "category": category,
                    "type": category,
                    "action": "REMOVED",
                    "value": identifier,
                }
            )

        for identifier in added[rename_count:]:
            changes.append(
                {
                    "category": category,
                    "type": category,
                    "action": "ADDED",
                    "value": identifier,
                }
            )

    add_rename_or_add_remove(
        "FUNCTION_IDENTIFIER",
        before["functions"],
        after["functions"],
    )

    add_rename_or_add_remove(
        "CLASS_IDENTIFIER",
        before["classes"],
        after["classes"],
    )

    # Arguments, variables and imports are reported as additions/removals.
    for category, key in (
        ("ARGUMENT_IDENTIFIER", "arguments"),
        ("VARIABLE_IDENTIFIER", "variables"),
        ("IMPORT_IDENTIFIER", "imports"),
    ):
        before_counter = Counter(before[key])
        after_counter = Counter(after[key])

        for identifier, count in (
            before_counter - after_counter
        ).items():
            for _ in range(count):
                changes.append(
                    {
                        "category": category,
                        "type": category,
                        "action": "REMOVED",
                        "value": identifier,
                    }
                )

        for identifier, count in (
            after_counter - before_counter
        ).items():
            for _ in range(count):
                changes.append(
                    {
                        "category": category,
                        "type": category,
                        "action": "ADDED",
                        "value": identifier,
                    }
                )

    return changes


# ============================================================
# BEHAVIOUR ANALYSIS
# ============================================================

def detect_behavior_categories(
    before_tree: ast.AST,
    after_tree: ast.AST,
    identifier_changes: list[dict[str, Any]],
) -> set[str]:
    """Detect potentially behavior-changing AST categories."""
    categories: set[str] = set()

    for change in identifier_changes:
        change_type = change.get("type", "")
        if change_type:
            categories.add(change_type)

    before_types = Counter(
        type(node).__name__
        for node in ast.walk(before_tree)
    )

    after_types = Counter(
        type(node).__name__
        for node in ast.walk(after_tree)
    )

    def count_changed(*node_types: str) -> bool:
        return any(
            before_types[node_type] != after_types[node_type]
            for node_type in node_types
        )

    if count_changed(
        "If",
        "IfExp",
        "Match",
        "match_case",
    ):
        categories.add("CONDITIONAL_LOGIC")

    if count_changed(
        "For",
        "AsyncFor",
        "While",
    ):
        categories.add("LOOP_LOGIC")

    if count_changed(
        "Break",
        "Continue",
        "Return",
        "Raise",
        "Try",
    ):
        categories.add("CONTROL_FLOW")

    if before_types["Call"] != after_types["Call"]:
        categories.add("FUNCTION_CALL")

    if count_changed(
        "BoolOp",
        "BinOp",
        "UnaryOp",
        "Compare",
        "IfExp",
        "NamedExpr",
    ):
        categories.add("EXPRESSION_LOGIC")

    if count_changed(
        "Assign",
        "AnnAssign",
        "AugAssign",
        "NamedExpr",
        "Delete",
    ):
        categories.add("ASSIGNMENT")

    if count_changed(
        "Subscript",
        "Attribute",
    ):
        categories.add("DATA_ACCESS")

    if count_changed(
        "With",
        "AsyncWith",
    ):
        categories.add("RESOURCE_HANDLING")

    if count_changed(
        "Import",
        "ImportFrom",
    ):
        categories.add("IMPORT")

    if count_changed(
        "FunctionDef",
        "AsyncFunctionDef",
        "ClassDef",
    ):
        categories.add("DEFINITION")

    if before_types["Constant"] != after_types["Constant"]:
        categories.add("LITERAL_VALUE")

    if count_changed(
        "Yield",
        "YieldFrom",
    ):
        categories.add("CONTROL_FLOW")

    if count_changed(
        "Global",
        "Nonlocal",
    ):
        categories.add("DATA_ACCESS")

    return categories


def classify_behavior_risk(
    categories: set[str],
    source_changed: bool,
) -> str:
    """Classify possible source-modification behavior risk."""
    if not source_changed:
        return "NONE"

    high_risk = {
        "CONDITIONAL_LOGIC",
        "LOOP_LOGIC",
        "EXPRESSION_LOGIC",
        "FUNCTION_CALL",
        "IMPORT",
        "ASSIGNMENT",
        "CONTROL_FLOW",
        "DATA_ACCESS",
        "RESOURCE_HANDLING",
        "DEFINITION",
        "FUNCTION_IDENTIFIER",
        "CLASS_IDENTIFIER",
        "LITERAL_VALUE",
    }

    medium_risk = {
        "VARIABLE_IDENTIFIER",
        "ARGUMENT_IDENTIFIER",
        "IMPORT_IDENTIFIER",
    }

    if categories & high_risk:
        return "HIGH"

    if categories & medium_risk:
        return "MEDIUM"

    return "LOW"


def analyze_source_behavior(
    before_source: str,
    after_source: str,
    file_name: str,
) -> dict[str, Any]:
    """Compare Python source before and after modification."""
    if before_source == after_source:
        return {
            "file": file_name,
            "changed": False,
            "syntax_valid": True,
            "ast_changed": False,
            "identifier_changes": [],
            "categories": set(),
            "risk": "NONE",
            "diff": [],
        }

    try:
        before_tree = parse_python_source(
            before_source,
            file_name,
        )
        after_tree = parse_python_source(
            after_source,
            file_name,
        )

    except SyntaxError:
        return {
            "file": file_name,
            "changed": True,
            "syntax_valid": False,
            "ast_changed": True,
            "identifier_changes": [],
            "categories": {"SYNTAX"},
            "risk": "HIGH",
            "diff": list(
                difflib.unified_diff(
                    before_source.splitlines(),
                    after_source.splitlines(),
                    fromfile=f"{file_name} (before)",
                    tofile=f"{file_name} (after)",
                    lineterm="",
                )
            ),
        }

    ast_changed = (
        normalize_ast(before_tree)
        != normalize_ast(after_tree)
    )

    identifier_changes = get_identifier_changes(
        before_tree,
        after_tree,
    )

    categories = detect_behavior_categories(
        before_tree,
        after_tree,
        identifier_changes,
    )

    risk = classify_behavior_risk(
        categories,
        source_changed=True,
    )

    diff = list(
        difflib.unified_diff(
            before_source.splitlines(),
            after_source.splitlines(),
            fromfile=f"{file_name} (before)",
            tofile=f"{file_name} (after)",
            lineterm="",
        )
    )

    return {
        "file": file_name,
        "changed": True,
        "syntax_valid": True,
        "ast_changed": ast_changed,
        "identifier_changes": identifier_changes,
        "categories": categories,
        "risk": risk,
        "diff": diff,
    }


def collect_behavior_reports(
    backup_dir: Path,
    source_files: list[Path],
) -> list[dict[str, Any]]:
    """Analyze every source file that changed after propagation."""
    reports: list[dict[str, Any]] = []

    for source_file in source_files:
        relative_path = source_file.relative_to(PROJECT_ROOT)
        backup_file = backup_dir / relative_path

        if not backup_file.exists():
            continue

        try:
            before_source = backup_file.read_text(
                encoding="utf-8"
            )
            after_source = source_file.read_text(
                encoding="utf-8"
            )
        except UnicodeDecodeError:
            continue

        if before_source == after_source:
            continue

        reports.append(
            analyze_source_behavior(
                before_source,
                after_source,
                str(relative_path),
            )
        )

    return reports


def print_behavior_reports(
    reports: list[dict[str, Any]],
) -> None:
    """Print source-change behavior analysis."""
    if not reports:
        print()
        print("Behavior analysis: no source changes detected.")
        return

    print()
    print("=" * 70)
    print("BEHAVIOR CHANGE ANALYSIS")
    print("=" * 70)

    for report in reports:
        print()
        print(f"FILE: {report['file']}")
        print(f"RISK LEVEL: {report['risk']}")

        if not report["syntax_valid"]:
            print("  Syntax validation: FAILED")
            continue

        if report["identifier_changes"]:
            print()
            print("  IDENTIFIER CHANGES:")

            for change in report["identifier_changes"]:
                print(
                    f"    {change['category']}: "
                    f"{change['value']}"
                )

        categories = sorted(report["categories"])

        if categories:
            print()
            print("  POTENTIAL BEHAVIOR CATEGORIES:")

            for category in categories:
                print(f"    - {category}")

        print()
        print(
            "  AST structure: "
            + (
                "CHANGED."
                if report["ast_changed"]
                else "UNCHANGED."
            )
        )

        print()
        print("  SOURCE DIFFERENCES:")

        for line in report["diff"][:40]:
            print(f"    {line}")

        if len(report["diff"]) > 40:
            print(
                f"    ... {len(report['diff']) - 40} "
                "additional diff lines"
            )

    print()
    print("=" * 70)


def behavior_change_requires_confirmation(
    reports: list[dict[str, Any]],
) -> bool:
    """Return True when source changes require operator confirmation."""
    return any(
        report["changed"]
        and (
            report["risk"] in {"MEDIUM", "HIGH"}
            or report["ast_changed"]
        )
        for report in reports
    )


def ask_keep_changes(
    reports: list[dict[str, Any]],
) -> bool:
    """Ask the operator whether potentially risky changes should remain."""
    if not behavior_change_requires_confirmation(reports):
        return True

    print()
    print("=" * 70)
    print("BEHAVIOR CHANGE WARNING")
    print("=" * 70)
    print()
    print(
        "The requested source modification may "
        "change program behaviour."
    )
    print(
        "Python syntax validity does NOT guarantee "
        "identical runtime behaviour."
    )
    print()
    print("Affected files:")

    for report in reports:
        if report["changed"]:
            print(
                f"  - {report['file']} "
                f"[{report['risk']} RISK]"
            )

    print()
    print(
        "A backup was created before source modification."
    )
    print()
    print("Do you want to KEEP these changes?")
    print("Y = Keep changes")
    print("N = Restore backup and rollback")

    while True:
        answer = input("Keep changes? [Y/N]: ").strip().lower()

        if answer == "y":
            return True

        if answer == "n":
            return False

        print("Please enter Y or N.")


# ============================================================
# MANUAL BACKUP RESTORATION
# ============================================================

def restore_latest_backup() -> bool:
    """Restore the most recent source-code backup."""
    if not BACKUP_DIRECTORY.exists():
        print("No backup directory exists.")
        return False

    backup_directories = sorted(
        (
            path
            for path in BACKUP_DIRECTORY.iterdir()
            if path.is_dir()
        ),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )

    if not backup_directories:
        print("No backups available.")
        return False

    latest_backup = backup_directories[0]

    print()
    print("=" * 70)
    print("BACKUP RESTORATION")
    print("=" * 70)
    print()
    print(
        f"Latest backup: "
        f"{latest_backup.relative_to(PROJECT_ROOT)}"
    )

    confirmation = input(
        "Restore this backup? [Y/N]: "
    ).strip().lower()

    if confirmation != "y":
        print("Restoration cancelled.")
        return False

    restore_backup(latest_backup)

    valid, errors = validate_python_sources()

    if not valid:
        print()
        print("Restored source failed Python validation.")

        for error in errors:
            print(f"  ERROR: {error}")

        return False

    print()
    print("Backup restoration completed successfully.")
    print("Python syntax validation: PASSED.")
    return True


# ============================================================
# SOURCE PROPAGATION
# ============================================================

def propagate_changes() -> None:
    """
    Authenticate and physically apply configured source changes.

    This function never modifies dataset files.
    """
    if not authenticate():
        print()
        print("Access denied. No source files were modified.")
        return

    source_files = get_source_files()

    if not source_files:
        print()
        print("No eligible Python source files found.")
        return

    print()
    print("=" * 70)
    print("SOURCE CODE PROPAGATION")
    print("=" * 70)

    # --------------------------------------------------------
    # BACKUP FIRST
    # --------------------------------------------------------

    backup_dir = create_backup(source_files)

    print()
    print(
        f"Backup created: "
        f"{backup_dir.relative_to(PROJECT_ROOT)}"
    )

    modifications = 0
    failed = False

    # --------------------------------------------------------
    # GLOBAL CHANGES
    # --------------------------------------------------------

    for change in GLOBAL_CHANGES:
        old_value = change.get("old")
        new_value = change.get("new")

        if not isinstance(old_value, str):
            print()
            print("[GLOBAL CHANGE]")
            print("FAILED: 'old' must be a string.")
            failed = True
            continue

        if not isinstance(new_value, str):
            print()
            print("[GLOBAL CHANGE]")
            print("FAILED: 'new' must be a string.")
            failed = True
            continue

        print()
        print("[GLOBAL CHANGE]")
        print(f"OLD: {old_value}")
        print(f"NEW: {new_value}")

        changed_files = replace_global(
            old_value,
            new_value,
        )

        if changed_files:
            for source_file in changed_files:
                print(
                    "  MODIFIED:",
                    source_file.relative_to(PROJECT_ROOT),
                )

            modifications += len(changed_files)
        else:
            print(
                "  No eligible Python source file "
                "contained the old value."
            )

    # --------------------------------------------------------
    # TARGETED CHANGES
    # --------------------------------------------------------

    for change in TARGETED_CHANGES:
        file_name = change.get("file")
        old_value = change.get("old")
        new_value = change.get("new")
        occurrence = change.get("occurrence")
        line_number = change.get("line")

        print()
        print("[TARGETED CHANGE]")
        print(f"FILE: {file_name}")
        print(f"OLD: {old_value}")
        print(f"NEW: {new_value}")
        print(f"OCCURRENCE: {occurrence}")
        print(f"LINE: {line_number}")

        if not isinstance(file_name, str):
            print("  FAILED: 'file' must be a string.")
            failed = True
            continue

        if not isinstance(old_value, str):
            print("  FAILED: 'old' must be a string.")
            failed = True
            continue

        if not isinstance(new_value, str):
            print("  FAILED: 'new' must be a string.")
            failed = True
            continue

        success, message = replace_targeted(
            file_name=file_name,
            old_value=old_value,
            new_value=new_value,
            occurrence=occurrence,
            line_number=line_number,
        )

        if success:
            print(f"  MODIFIED: {message}")
            modifications += 1
        else:
            print(f"  FAILED: {message}")
            failed = True

    # --------------------------------------------------------
    # SYNTAX VALIDATION
    # --------------------------------------------------------

    print()
    print("Validating Python source files...")

    valid, errors = validate_python_sources()

    if not valid:
        print()
        print("VALIDATION FAILED.")
        print("Restoring backup...")

        restore_backup(backup_dir)

        print("Rollback completed.")

        for error in errors:
            print(f"  ERROR: {error}")

        return

    print("Python syntax validation: PASSED.")

    # --------------------------------------------------------
    # BEHAVIOR ANALYSIS
    # --------------------------------------------------------

    behavior_reports = collect_behavior_reports(
        backup_dir=backup_dir,
        source_files=source_files,
    )

    print_behavior_reports(behavior_reports)

    # --------------------------------------------------------
    # USER APPROVAL
    # --------------------------------------------------------

    if behavior_change_requires_confirmation(
        behavior_reports
    ):
        keep_changes = ask_keep_changes(
            behavior_reports
        )

        if not keep_changes:
            print()
            print("Changes rejected by user.")
            print("Restoring backup...")

            restore_backup(backup_dir)

            print("Rollback completed.")
            return

        print()
        print(
            "User approved the potential "
            "behavior changes."
        )

    # --------------------------------------------------------
    # FINAL RESULT
    # --------------------------------------------------------

    print()
    print("=" * 70)
    print("SOURCE PROPAGATION COMPLETED")
    print("=" * 70)
    print()
    print(
        f"Successful modifications: {modifications}"
    )

    if failed:
        print("Some requested modifications failed.")
    else:
        print("All requested modifications processed.")

    print()
    print(
        "Changes were physically written to the "
        "eligible Python source files."
    )


# ============================================================
# ENVIRONMENT INFORMATION
# ============================================================

def print_environment() -> None:
    """Display centralized environment configuration."""
    print()
    print("=" * 70)
    print(CURRENT_VALUES["PROJECT_NAME"])
    print("ENVIRONMENT CONFIGURATION")
    print("=" * 70)
    print()
    print(f"Project Root       : {PROJECT_ROOT}")
    print(f"Raw Data Directory : {RAW_DATA_DIRECTORY}")
    print(f"Processed Directory: {PROCESSED_DATA_DIRECTORY}")
    print(f"Input Directory    : {INPUT_DATA_DIRECTORY}")
    print(f"Company Data       : {COMPANY_DATA_DIRECTORY}")
    print(f"Output Directory   : {OUTPUTS_DIRECTORY}")
    print(f"Power BI Directory : {POWERBI_DIRECTORY}")
    print()
    print(f"Active Dataset     : {ACTIVE_DATASET_PATH}")
    print(
        f"Dataset Format     : "
        f"{CURRENT_VALUES['ACTIVE_DATASET_FORMAT']}"
    )
    print(
        f"Maximum Companies  : "
        f"{CURRENT_VALUES['MAX_COMPANIES_PER_DATASET']}"
    )
    print(
        f"Maximum File Size  : "
        f"{CURRENT_VALUES['MAX_DATASET_FILE_SIZE_MB']} MB"
    )
    print()
    print(
        "Dataset architecture: "
        "runtime-discovered and dataset-independent."
    )
    print()
    print("Environment paths loaded successfully.")


# ============================================================
# ENTRY POINT
# ============================================================

def print_usage() -> None:
    """Print supported command-line operations."""
    print("Available commands:")
    print("  python env.py")
    print("  python env.py --restore")
    print("  python env.py --paths")
    print("  python env.py --init")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        command = sys.argv[1].lower()

        if command == "--restore":
            if not authenticate():
                print()
                print(
                    "Access denied. "
                    "Backup was not restored."
                )
            else:
                restore_latest_backup()

        elif command == "--paths":
            print_environment()

        elif command == "--init":
            ensure_project_directories()
            print()
            print(
                "Required project directories "
                "initialized successfully."
            )
            print_environment()

        else:
            print(f"Unknown command: {sys.argv[1]}")
            print()
            print_usage()

    else:
        propagate_changes()

