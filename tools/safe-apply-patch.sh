#!/usr/bin/env bash
# Safe patch application script for auto-improver
# This script safely applies patches with validation and rollback capability

set -euo pipefail

PATCH_FILE="$1"
REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

if [[ ! -f "$PATCH_FILE" ]]; then
    echo "Error: Patch file '$PATCH_FILE' not found"
    exit 1
fi

echo "Applying patch: $PATCH_FILE"

# Create backup
BACKUP_DIR="$REPO_ROOT/backups/$(date +%Y%m%d_%H%M%S)"
mkdir -p "$BACKUP_DIR"

# Backup modified files
echo "Creating backup in $BACKUP_DIR"
git diff --name-only | while read -r file; do
    if [[ -f "$file" ]]; then
        mkdir -p "$(dirname "$BACKUP_DIR/$file")"
        cp "$file" "$BACKUP_DIR/$file"
    fi
done

# Apply patch
if git apply --check "$PATCH_FILE" 2>/dev/null; then
    echo "Patch validation passed, applying..."
    if git apply "$PATCH_FILE"; then
        echo "Patch applied successfully"
        
        # Run tests to validate
        echo "Running validation tests..."
        if cd "$REPO_ROOT" && python3 -m pytest tests/ -v --tb=short; then
            echo "✅ Patch applied and validated successfully"
            exit 0
        else
            echo "❌ Tests failed after patch application, rolling back..."
            git checkout -- .
            exit 1
        fi
    else
        echo "❌ Failed to apply patch"
        exit 1
    fi
else
    echo "❌ Patch validation failed"
    exit 1
fi
