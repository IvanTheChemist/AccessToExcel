from access_to_excel.exporter import sanitize_sheet_name, dedupe_names


def test_sanitize_basic():
    assert sanitize_sheet_name("Customers") == "Customers"
    assert sanitize_sheet_name("[Bad]*Name?") == "_Bad__Name_"


def test_length_limit():
    long = "x" * 100
    assert len(sanitize_sheet_name(long)) == 31


def test_dedupe():
    names = ["Orders", "Orders", "Orders:2024", "Orders:2024"]
    deduped = dedupe_names(names)
    assert deduped[0] == "Orders"
    assert deduped[1].startswith("Orders_")
    assert len(set(deduped)) == len(deduped)
