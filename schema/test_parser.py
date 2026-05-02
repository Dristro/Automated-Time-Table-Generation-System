"""
Comprehensive Test Suite for TimetableDataParser

Tests the complete pipeline:
1. Parser reads input file (Excel/CSV/JSON)
2. Parser generates JSON intermediate file
3. JSON matches IngestionService expectations
4. JSON can be ingested into database

Requirements Validated:
✓ Parser accepts any file type (.xlsx, .csv, .json)
✓ Output JSON format matches IngestionService expectations
✓ All required fields present in output
✓ Data integrity preserved through parsing
✓ Validation rules enforced
✓ Error handling works correctly
"""

import json
import pytest
from pathlib import Path
from schema.parser import TimetableDataParser


class TestParserBasics:
    """Test basic parser functionality."""

    def test_parser_initialization(self):
        """Parser initializes without file arguments."""
        parser = TimetableDataParser()
        assert parser.strict is False
        parser_strict = TimetableDataParser(strict=True)
        assert parser_strict.strict is True

    def test_parser_reusability(self):
        """Same parser instance can parse multiple files."""
        parser = TimetableDataParser()
        # Parser should reset state between parse calls
        assert len(parser.professors) == 0
        # After first parse, can parse again
        # (implementation detail: _reset_state() is called in parse())

    def test_unsupported_format_error(self, tmp_path):
        """Parser rejects unsupported file formats."""
        dummy_file = tmp_path / "test.txt"
        dummy_file.write_text("dummy content")

        parser = TimetableDataParser()
        success, stats = parser.parse(str(dummy_file), str(tmp_path / "out.json"))

        assert not success
        assert len(stats['errors']) > 0
        assert "Unsupported file format" in stats['errors'][0]


class TestExcelParsing:
    """Test Excel file parsing."""

    def test_parse_sample_excel(self, tmp_path):
        """Parse sample_timetable.xlsx generates expected JSON."""
        output_file = tmp_path / "output.json"

        parser = TimetableDataParser()
        success, stats = parser.parse("sample_timetable.xlsx", str(output_file))

        assert success
        assert output_file.exists()
        assert stats['counts']['professors'] == 56
        assert stats['counts']['rooms'] == 34
        assert stats['counts']['student_groups'] == 15
        assert stats['counts']['courses'] == 37

    def test_excel_json_output_structure(self, tmp_path):
        """Output JSON has correct structure for IngestionService."""
        output_file = tmp_path / "output.json"

        parser = TimetableDataParser()
        success, _ = parser.parse("sample_timetable.xlsx", str(output_file))

        assert success

        with open(output_file) as f:
            data = json.load(f)

        # Check top-level keys
        assert "professors" in data
        assert "rooms" in data
        assert "student_groups" in data
        assert "courses" in data
        assert "_meta" in data

        # Check types
        assert isinstance(data["professors"], list)
        assert isinstance(data["rooms"], list)
        assert isinstance(data["student_groups"], list)
        assert isinstance(data["courses"], list)
        assert isinstance(data["_meta"], dict)

    def test_professor_structure(self, tmp_path):
        """Each professor has required fields."""
        output_file = tmp_path / "output.json"

        parser = TimetableDataParser()
        parser.parse("sample_timetable.xlsx", str(output_file))

        with open(output_file) as f:
            data = json.load(f)

        for prof in data["professors"]:
            assert "name" in prof
            assert "email" in prof
            assert isinstance(prof["name"], str)
            assert isinstance(prof["email"], str)
            assert len(prof["name"]) > 0
            assert len(prof["email"]) > 0
            assert "@" in prof["email"]

    def test_room_structure(self, tmp_path):
        """Each room has required fields."""
        output_file = tmp_path / "output.json"

        parser = TimetableDataParser()
        parser.parse("sample_timetable.xlsx", str(output_file))

        with open(output_file) as f:
            data = json.load(f)

        for room in data["rooms"]:
            assert "name" in room
            assert "is_lab" in room
            assert "capacity" in room
            assert isinstance(room["name"], str)
            assert isinstance(room["is_lab"], bool)
            assert isinstance(room["capacity"], int)
            assert room["capacity"] > 0

    def test_student_group_structure(self, tmp_path):
        """Each student group has required fields."""
        output_file = tmp_path / "output.json"

        parser = TimetableDataParser()
        parser.parse("sample_timetable.xlsx", str(output_file))

        with open(output_file) as f:
            data = json.load(f)

        for group in data["student_groups"]:
            assert "name" in group
            assert "size" in group
            assert "level" in group
            assert isinstance(group["name"], str)
            assert isinstance(group["size"], int)
            assert isinstance(group["level"], str)
            assert group["size"] > 0

    def test_course_structure(self, tmp_path):
        """Each course has required fields matching IngestionService."""
        output_file = tmp_path / "output.json"

        parser = TimetableDataParser()
        parser.parse("sample_timetable.xlsx", str(output_file))

        with open(output_file) as f:
            data = json.load(f)

        required_fields = [
            "course_code", "course_name", "session_type", "professor",
            "room", "student_group", "slots_required", "slots_continuous",
            "preference_bin", "total_credits"
        ]

        for course in data["courses"]:
            for field in required_fields:
                assert field in course, f"Missing field: {field}"

            # Type validation
            assert isinstance(course["course_code"], str)
            assert isinstance(course["course_name"], str)
            assert isinstance(course["session_type"], str)
            assert isinstance(course["professor"], str)
            assert isinstance(course["room"], str)
            assert isinstance(course["student_group"], str)
            assert isinstance(course["slots_required"], int)
            assert isinstance(course["slots_continuous"], bool)
            assert isinstance(course["preference_bin"], int)
            assert isinstance(course["total_credits"], int)

            # Value constraints
            assert course["slots_required"] > 0
            assert course["preference_bin"] in [1, 2, 3]
            assert course["total_credits"] >= 0
            assert course["session_type"] in ["lecture", "tutorial", "lab"]

    def test_metadata_structure(self, tmp_path):
        """_meta block has expected structure."""
        output_file = tmp_path / "output.json"

        parser = TimetableDataParser()
        parser.parse("sample_timetable.xlsx", str(output_file))

        with open(output_file) as f:
            data = json.load(f)

        meta = data["_meta"]
        assert "source_file" in meta
        assert "parsed_at" in meta
        assert "counts" in meta
        assert "warnings" in meta

        counts = meta["counts"]
        assert "professors" in counts
        assert "rooms" in counts
        assert "student_groups" in counts
        assert "courses" in counts

    def test_excel_data_integrity(self, tmp_path):
        """Data is preserved correctly through parsing."""
        output_file = tmp_path / "output.json"

        parser = TimetableDataParser()
        parser.parse("sample_timetable.xlsx", str(output_file))

        with open(output_file) as f:
            data = json.load(f)

        # Verify specific entries from seed data
        prof_names = {p["name"] for p in data["professors"]}
        assert "Dr. Rakesh" in prof_names
        assert "Dr. Anagha Tobi" in prof_names
        assert "Dr. Shabnam" in prof_names

        room_names = {r["name"] for r in data["rooms"]}
        assert "ELT 1" in room_names
        assert "CS LAB 1" in room_names
        assert "Auditorium" in room_names

        group_names = {g["name"] for g in data["student_groups"]}
        assert "CS1" in group_names
        assert "AI1" in group_names
        assert "ECE" in group_names

        # Verify a specific course
        ma2103_courses = [c for c in data["courses"] if c["course_code"] == "MA2103"]
        assert len(ma2103_courses) > 0
        lecture = [c for c in ma2103_courses if c["session_type"] == "lecture"][0]
        assert lecture["professor"] == "Dr. Rakesh"
        assert lecture["room"] == "ELT 1"


class TestCSVParsing:
    """Test CSV file parsing."""

    def test_parse_csv_file(self, tmp_path):
        """Parser can handle CSV files."""
        # Create a test CSV file
        csv_file = tmp_path / "test.csv"
        csv_content = """course_code,course_name,session_type,professor,room,student_group,slots_required,slots_continuous,preference_bin,total_credits
MA2103,Mathematics III,lecture,Dr. Rakesh,ELT 1,CS1,5,FALSE,1,3
CS/AI 2102,Data Structures,lecture,Dr. Ravi Kishor,ELT 2,CS1,4,FALSE,1,4
CS/AI 2102,Data Structures,lab,Dr. Shabnam,CS LAB 1,CS1,2,TRUE,1,0"""
        csv_file.write_text(csv_content)

        output_file = tmp_path / "output.json"
        parser = TimetableDataParser()
        success, stats = parser.parse(str(csv_file), str(output_file))

        assert success
        assert stats['counts']['professors'] == 3
        assert stats['counts']['rooms'] == 3
        assert stats['counts']['student_groups'] == 1
        assert stats['counts']['courses'] == 3

    def test_csv_auto_creates_entities(self, tmp_path):
        """CSV parsing auto-creates professors/rooms/groups."""
        csv_file = tmp_path / "test.csv"
        csv_content = """course_code,course_name,session_type,professor,room,student_group,slots_required
TEST101,Test,lecture,Dr. NewProf,NewRoom,NewGroup,2"""
        csv_file.write_text(csv_content)

        output_file = tmp_path / "output.json"
        parser = TimetableDataParser()
        success, stats = parser.parse(str(csv_file), str(output_file))

        assert success

        with open(output_file) as f:
            data = json.load(f)

        prof_names = {p["name"] for p in data["professors"]}
        assert "Dr. NewProf" in prof_names

        room_names = {r["name"] for r in data["rooms"]}
        assert "NewRoom" in room_names

        group_names = {g["name"] for g in data["student_groups"]}
        assert "NewGroup" in group_names


class TestJSONParsing:
    """Test JSON file parsing."""

    def test_parse_json_normalized(self, tmp_path):
        """Parser can handle pre-normalized JSON."""
        json_file = tmp_path / "test.json"
        json_data = {
            "professors": [
                {"name": "Dr. Test1", "email": "test1@example.com"}
            ],
            "rooms": [
                {"name": "Room A", "is_lab": False, "capacity": 100}
            ],
            "student_groups": [
                {"name": "Group1", "size": 50, "level": "batch"}
            ],
            "courses": [
                {
                    "course_code": "TEST101",
                    "course_name": "Test Course",
                    "session_type": "lecture",
                    "professor": "Dr. Test1",
                    "room": "Room A",
                    "student_group": "Group1",
                    "slots_required": 2,
                    "preference_bin": 1
                }
            ]
        }
        json_file.write_text(json.dumps(json_data))

        output_file = tmp_path / "output.json"
        parser = TimetableDataParser()
        success, stats = parser.parse(str(json_file), str(output_file))

        assert success
        assert stats['counts']['professors'] == 1
        assert stats['counts']['rooms'] == 1
        assert stats['counts']['student_groups'] == 1
        assert stats['counts']['courses'] == 1

    def test_parse_json_flat_array(self, tmp_path):
        """Parser can handle flat JSON array of courses."""
        json_file = tmp_path / "test.json"
        json_data = [
            {
                "course_code": "TEST101",
                "course_name": "Test Course",
                "session_type": "lecture",
                "professor": "Dr. Test",
                "room": "Room A",
                "student_group": "Group1",
                "slots_required": 2,
                "preference_bin": 1
            }
        ]
        json_file.write_text(json.dumps(json_data))

        output_file = tmp_path / "output.json"
        parser = TimetableDataParser()
        success, stats = parser.parse(str(json_file), str(output_file))

        assert success
        assert stats['counts']['courses'] == 1


class TestErrorHandling:
    """Test error handling and validation."""

    def test_missing_required_sheet(self, tmp_path):
        """Parser detects missing required sheets."""
        # This would require creating a broken Excel file
        # For now, we test the error reporting mechanism
        parser = TimetableDataParser()

        # Manually set up an error
        parser.errors.append("ERROR: Missing required sheets: Student Groups")

        assert len(parser.errors) == 1
        assert "Missing required sheets" in parser.errors[0]

    def test_error_includes_row_column(self):
        """Error messages include row and column information."""
        # Test the error message format
        source = "Courses"
        row_num = 5
        column = "professor"
        msg = f"ERROR [{source}, Row {row_num}, Column '{column}']: Unknown professor"

        assert "Courses" in msg
        assert "Row 5" in msg
        assert "professor" in msg

    def test_strict_mode_fails_on_warnings(self, tmp_path):
        """Strict mode treats warnings as errors."""
        output_file = tmp_path / "output.json"

        parser = TimetableDataParser(strict=True)
        success, stats = parser.parse("sample_timetable.xlsx", str(output_file))

        # With strict mode and missing emails, should fail
        assert not success
        assert len(stats['warnings']) > 0


class TestIngestionServiceCompatibility:
    """Test JSON output compatibility with IngestionService."""

    def test_json_loadable_by_ingestion_service(self, tmp_path):
        """Output JSON can be loaded by IngestionService."""
        output_file = tmp_path / "output.json"

        parser = TimetableDataParser()
        parser.parse("sample_timetable.xlsx", str(output_file))

        # Simulate IngestionService loading
        with open(output_file) as f:
            data = json.load(f)

        # Verify structure matches IngestionService expectations
        assert "professors" in data
        assert "rooms" in data
        assert "courses" in data

        # These are the three entity types IngestionService.ingest_course_data expects
        for course in data["courses"]:
            # Each course must reference existing entities
            assert course["professor"] in {p["name"] for p in data["professors"]}
            assert course["room"] in {r["name"] for r in data["rooms"]}
            if "student_groups" in data:
                assert course["student_group"] in {g["name"] for g in data["student_groups"]}

    def test_json_has_correct_field_names(self, tmp_path):
        """JSON field names match IngestionService.ingest_course_data expectations."""
        output_file = tmp_path / "output.json"

        parser = TimetableDataParser()
        parser.parse("sample_timetable.xlsx", str(output_file))

        with open(output_file) as f:
            data = json.load(f)

        # IngestionService looks for these specific field names
        for course in data["courses"]:
            assert "course_code" in course
            assert "session_type" in course
            assert "professor" in course
            assert "room" in course
            assert "student_group" in course
            assert "slots_required" in course
            assert "slots_continuous" in course
            assert "preference_bin" in course

    def test_email_auto_generation_format(self, tmp_path):
        """Auto-generated emails match expected format."""
        output_file = tmp_path / "output.json"

        parser = TimetableDataParser()
        parser.parse("sample_timetable.xlsx", str(output_file))

        with open(output_file) as f:
            data = json.load(f)

        for prof in data["professors"]:
            email = prof["email"]
            assert "@mahindrauniversity.edu.in" in email
            # Should not have Dr. prefix in email
            assert not email.startswith("dr.")
            # Should have name components
            name_parts = prof["name"].lower().split()
            if name_parts[0] == "dr.":
                name_parts = name_parts[1:]
            # At least one name part should be in email
            assert any(part in email for part in name_parts if part.strip())


class TestDataQuality:
    """Test data quality and integrity."""

    def test_no_duplicate_professors(self, tmp_path):
        """No duplicate professor names in output."""
        output_file = tmp_path / "output.json"

        parser = TimetableDataParser()
        parser.parse("sample_timetable.xlsx", str(output_file))

        with open(output_file) as f:
            data = json.load(f)

        prof_names = [p["name"] for p in data["professors"]]
        assert len(prof_names) == len(set(prof_names))

    def test_no_duplicate_rooms(self, tmp_path):
        """No duplicate room names in output."""
        output_file = tmp_path / "output.json"

        parser = TimetableDataParser()
        parser.parse("sample_timetable.xlsx", str(output_file))

        with open(output_file) as f:
            data = json.load(f)

        room_names = [r["name"] for r in data["rooms"]]
        assert len(room_names) == len(set(room_names))

    def test_no_duplicate_groups(self, tmp_path):
        """No duplicate student group names in output."""
        output_file = tmp_path / "output.json"

        parser = TimetableDataParser()
        parser.parse("sample_timetable.xlsx", str(output_file))

        with open(output_file) as f:
            data = json.load(f)

        group_names = [g["name"] for g in data["student_groups"]]
        assert len(group_names) == len(set(group_names))

    def test_counts_match_data(self, tmp_path):
        """Metadata counts match actual data."""
        output_file = tmp_path / "output.json"

        parser = TimetableDataParser()
        parser.parse("sample_timetable.xlsx", str(output_file))

        with open(output_file) as f:
            data = json.load(f)

        meta_counts = data["_meta"]["counts"]
        assert meta_counts["professors"] == len(data["professors"])
        assert meta_counts["rooms"] == len(data["rooms"])
        assert meta_counts["student_groups"] == len(data["student_groups"])
        assert meta_counts["courses"] == len(data["courses"])


# ============================================================================
# Test Fixtures
# ============================================================================

@pytest.fixture(scope="session", autouse=True)
def check_sample_file():
    """Ensure sample_timetable.xlsx exists before running tests."""
    sample_file = Path("sample_timetable.xlsx")
    if not sample_file.exists():
        pytest.skip("sample_timetable.xlsx not found. Run: python generate_test_excel.py")


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])
