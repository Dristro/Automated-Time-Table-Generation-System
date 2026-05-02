# Parser Refactoring — Design Changes

## Summary

The `TimetableDataParser` has been refactored from a **file-specific** implementation to a **file-agnostic** design that supports multiple input formats while maintaining a clean, reusable interface.

---

## Key Changes

### 1. Initialization (No File Arguments)

**Before:**
```python
parser = TimetableDataParser(input_path="data.xlsx", output_path="out.json", strict=False)
success, stats = parser.parse()
```

**After:**
```python
parser = TimetableDataParser(strict=False)
success, stats = parser.parse("data.xlsx", "out.json")
```

**Benefits:**
- Parser is created once, reused for multiple files
- File path and format are provided at parse time
- Cleaner instantiation; only configuration (strict mode) in __init__

### 2. Format Auto-Detection

**Before:**
- Only Excel (.xlsx) files supported
- Had to know the exact sheet structure

**After:**
- Supports 3 input formats: `.xlsx`, `.csv`, `.json`
- Format auto-detected from file extension
- Each format has its own loader method: `_parse_excel()`, `_parse_csv()`, `_parse_json()`

### 3. Internal State Management

**New Method: `_reset_state()`**
```python
def _reset_state(self) -> None:
    """Reset internal state for a new parse operation."""
    self.errors = []
    self.warnings = []
    self.professors = []
    self.rooms = []
    self.student_groups = []
    self.courses = []
    self.professor_names = set()
    self.room_names = set()
    self.student_group_names = set()
```

Called at the start of `parse()` so the same parser instance can handle multiple files without cross-contamination.

### 4. Shared Helper Methods

All format parsers (Excel, CSV, JSON) use the same validation and normalization methods:

```
_add_professor_from_dict()      # Used by all formats
_add_room_from_dict()           # Used by all formats
_add_student_group_from_dict()  # Used by all formats
_add_course_from_dict()         # Used by all formats
```

**Benefits:**
- Consistent validation across formats
- Validation logic is format-agnostic
- DRY principle: no duplicated validation code

### 5. Format-Specific Loaders

#### `_parse_excel()`
- Expects 4 sheets: Professors, Rooms, Student Groups, Courses
- Uses openpyxl to read sheets
- Calls Excel-specific `_parse_professors()`, `_parse_rooms()`, etc.

#### `_parse_csv()`
- Single CSV file with course instances
- Auto-creates professors/rooms/student groups if not present
- Uses Python's `csv` module (standard library)

#### `_parse_json()`
- Supports pre-normalized structure or flat array of courses
- Auto-expands flat arrays to full structure
- Uses Python's `json` module (standard library)

### 6. Dependency Management

openpyxl is now **optional**:
- If missing, Excel parsing fails with helpful error message
- CSV and JSON parsing still work
- Graceful degradation

```python
HAS_OPENPYXL = True  # or False if import fails
```

---

## Error Reporting Improvements

### Error Messages Include Format Agnostic Context

```python
# All formats report errors similarly:
ERROR [source, Row row_num, Column 'field']: error message
```

Where `source` can be:
- Sheet name (Excel): "Professors", "Rooms", etc.
- Format name (CSV/JSON): "CSV", "JSON"

### Example Error Flow

```python
# Excel source
ERROR [Professors, Row 5, Column 'name']: Duplicate professor name 'Dr. Unknown'

# CSV source
ERROR [CSV, Row 5, Column 'professor']: Unknown professor 'Dr. Unknown'

# JSON source
ERROR [JSON, Row 5, Column 'professor']: Unknown professor 'Dr. Unknown'
```

---

## Validation Strategy (Unchanged)

### Phase 1: Per-Source Validation
- Required fields present
- Type constraints (bool, int, enum)
- Uniqueness within a source
- Duplicates detected

### Phase 2: Cross-Source Validation
- Professor references exist
- Room references exist
- Student group references exist

Both phases report ALL errors before exiting (not fail-fast).

---

## CLI Changes

### Updated Help Text
```
$ python schema/parser.py --help
usage: parser.py [-h] --input INPUT [--output OUTPUT] [--strict]

Parse timetable data (.xlsx, .csv, .json) and produce JSON intermediate format.

options:
  -h, --help           show this help message and exit
  --input, -i INPUT    Path to input file (.xlsx, .csv, or .json)
  --output, -o OUTPUT  Path to output JSON file (default: schema/output/parsed_data.json)
  --strict             Treat warnings as errors and fail if any are found
```

### Examples
```bash
# Excel
python schema/parser.py --input data.xlsx --output parsed.json

# CSV
python schema/parser.py --input courses.csv --output parsed.json

# JSON
python schema/parser.py --input seed.json --output normalized.json
```

---

## Architectural Benefits

1. **Reusability**: Same parser instance handles multiple files
2. **Format Agnostic**: Validation logic doesn't depend on input format
3. **Extensibility**: Adding a new format (e.g., `.ods`, `.xls`) is straightforward:
   - Add `_parse_ods()` method
   - Wire it in `parse()` with new file extension check
   - Uses same `_add_*_from_dict()` helpers

4. **Separation of Concerns**: 
   - Format loading (Excel/CSV/JSON specific)
   - Data normalization (format-agnostic)
   - Validation (format-agnostic)
   - Output (format-agnostic)

5. **Testability**: Can test validation independently of file format

---

## Backward Compatibility

The CLI remains fully backward compatible:
```bash
# Still works
python schema/parser.py --input data.xlsx --output parsed.json
```

Programmatic API changed intentionally to support file-agnostic design:
```python
# Old API (file-specific)
parser = TimetableDataParser(input_path="data.xlsx")
success, _ = parser.parse()

# New API (file-agnostic)
parser = TimetableDataParser()
success, _ = parser.parse("data.xlsx", "out.json")
```

---

## Testing Considerations

### Unit Tests
Can now test validation without loading files:
```python
parser = TimetableDataParser()
parser._add_professor_from_dict({"name": "Dr. Test"}, source="unit_test", row_num=1)
assert parser.professor_names == {"Dr. Test"}
```

### Integration Tests
Can test each format separately:
```python
# Test Excel parsing
parser.parse("test.xlsx", "out.json")

# Test CSV parsing
parser.parse("test.csv", "out.json")

# Test JSON parsing
parser.parse("test.json", "out.json")
```

---

## Future Enhancements

1. **Streaming for Large Files**: Replace list-based storage with generators
2. **Batch Validation**: Validate before normalizing (two-pass approach)
3. **Custom Validators**: Allow injection of format-specific validators
4. **Schema Inference**: Auto-detect column order in CSV files
5. **Parallel Parsing**: Use async/threading for large files
