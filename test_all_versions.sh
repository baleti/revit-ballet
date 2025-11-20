#!/bin/bash

# Test script for building Revit Ballet across all supported versions
# Usage: ./test_all_versions.sh [version1 version2 ...]
# If no versions specified, tests all versions from 2017-2026

set -e

cd "$(dirname "$0")"

# Color codes for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Default versions to test (all supported versions)
if [ $# -eq 0 ]; then
    VERSIONS=(2026 2025 2024 2023 2022 2021 2020 2019 2018 2017)
else
    VERSIONS=("$@")
fi

echo "========================================="
echo "Revit Ballet Multi-Version Build Test"
echo "========================================="
echo ""

SUCCESS_COUNT=0
FAIL_COUNT=0
FAILED_VERSIONS=()

for version in "${VERSIONS[@]}"; do
    echo "Testing Revit $version..."

    # Build the project
    if dotnet build commands/commands.csproj -c Release -p:RevitYear=$version > /tmp/build_$version.log 2>&1; then
        # Count errors in build output
        ERROR_COUNT=$(grep -c "error CS" /tmp/build_$version.log || true)

        if [ $ERROR_COUNT -eq 0 ]; then
            echo -e "  ${GREEN}✓${NC} Revit $version: BUILD SUCCESSFUL (0 errors)"
            ((SUCCESS_COUNT++))
        else
            echo -e "  ${RED}✗${NC} Revit $version: BUILD FAILED ($ERROR_COUNT errors)"
            echo "    See /tmp/build_$version.log for details"
            ((FAIL_COUNT++))
            FAILED_VERSIONS+=($version)
        fi
    else
        # Build command itself failed
        ERROR_COUNT=$(grep -c "error CS" /tmp/build_$version.log || true)
        echo -e "  ${RED}✗${NC} Revit $version: BUILD FAILED ($ERROR_COUNT errors)"
        echo "    See /tmp/build_$version.log for details"
        ((FAIL_COUNT++))
        FAILED_VERSIONS+=($version)
    fi

    echo ""
done

echo "========================================="
echo "Build Summary"
echo "========================================="
echo -e "${GREEN}Successful builds:${NC} $SUCCESS_COUNT"
echo -e "${RED}Failed builds:${NC} $FAIL_COUNT"

if [ $FAIL_COUNT -gt 0 ]; then
    echo ""
    echo "Failed versions: ${FAILED_VERSIONS[*]}"
    echo ""
    echo "To view detailed errors for a specific version, run:"
    echo "  cat /tmp/build_<version>.log | grep 'error CS'"
    exit 1
else
    echo ""
    echo -e "${GREEN}All builds successful!${NC}"
    exit 0
fi
