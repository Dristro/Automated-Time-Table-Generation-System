#!/bin/bash
# Comprehensive test runner for TimetableDataParser
#
# Requirements validation:
# 1. Parser reads input file (Excel/CSV/JSON)
# 2. Parser generates JSON intermediate file
# 3. JSON matches IngestionService expectations
# 4. All tests pass before integration

set -e

PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$PROJECT_ROOT"

# Activate virtual environment
if [ ! -d ".venv" ]; then
    echo "❌ Virtual environment not found at .venv"
    exit 1
fi

source .venv/bin/activate

# Check dependencies
echo "═══════════════════════════════════════════════════════════"
echo "Checking dependencies..."
echo "═══════════════════════════════════════════════════════════"

python -c "import openpyxl" && echo "✓ openpyxl installed" || (echo "❌ openpyxl not found"; pip install openpyxl -q; echo "✓ openpyxl installed")
python -c "import pytest" && echo "✓ pytest installed" || (echo "❌ pytest not found"; pip install pytest -q; echo "✓ pytest installed")

# Check test data
echo ""
echo "═══════════════════════════════════════════════════════════"
echo "Checking test data..."
echo "═══════════════════════════════════════════════════════════"

if [ ! -f "sample_timetable.xlsx" ]; then
    echo "⚠️  sample_timetable.xlsx not found. Generating..."
    python generate_test_excel.py
    echo "✓ Test data generated"
else
    echo "✓ sample_timetable.xlsx found"
fi

# Run parser syntax check
echo ""
echo "═══════════════════════════════════════════════════════════"
echo "Syntax validation..."
echo "═══════════════════════════════════════════════════════════"

python -m py_compile schema/parser.py
echo "✓ schema/parser.py syntax valid"

# Run comprehensive test suite
echo ""
echo "═══════════════════════════════════════════════════════════"
echo "Running comprehensive test suite..."
echo "═══════════════════════════════════════════════════════════"
echo ""

python -m pytest schema/test_parser.py -v --tb=short

# Summary
echo ""
echo "═══════════════════════════════════════════════════════════"
echo "✅ ALL TESTS PASSED"
echo "═══════════════════════════════════════════════════════════"
echo ""
echo "Test Coverage:"
echo "  • Parser Basics: initialization, reusability, error handling"
echo "  • Excel Parsing: structure, fields, data integrity"
echo "  • CSV Parsing: auto-creation of entities"
echo "  • JSON Parsing: normalized and flat formats"
echo "  • Error Handling: validation, row/column tracking"
echo "  • IngestionService Compatibility: format validation"
echo "  • Data Quality: no duplicates, counts match"
echo ""
echo "Next Steps:"
echo "  1. Verify parsed JSON with: cat parsed_sample.json | jq ."
echo "  2. Integrate with IngestionService: IngestionService.perform_initial_seeding('parsed_sample.json')"
echo "  3. Run database ingestion tests"
echo ""
