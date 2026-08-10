"""Backend tests for Pick Ticket / Batch Print History endpoints (Iteration 5)."""
import os
import time
import requests
import pytest

BASE_URL = os.environ.get("REACT_APP_BACKEND_URL", "https://sap-uploader.preview.emergentagent.com").rstrip("/")
API = f"{BASE_URL}/api"


@pytest.fixture(scope="module")
def session():
    s = requests.Session()
    s.headers.update({"Content-Type": "application/json"})
    return s


def test_root_health(session):
    r = session.get(f"{API}/")
    assert r.status_code == 200
    assert "message" in r.json()


def test_check_empty_returns_no_prints(session):
    r = session.post(f"{API}/pick-tickets/check", json={"to_numbers": []})
    assert r.status_code == 200
    assert r.json() == {"previous_prints": []}


def test_check_unknown_tos(session):
    r = session.post(f"{API}/pick-tickets/check", json={"to_numbers": ["TEST_UNKNOWN_1", "TEST_UNKNOWN_2"]})
    assert r.status_code == 200
    assert r.json().get("previous_prints") == []


def test_create_and_list_pick_ticket(session):
    payload = {
        "batch_code": "PTF-Batch-1-TESTX",
        "batch_number": 1,
        "batch_size": 3,
        "picker_count": 3,
        "to_numbers": ["TEST_5330150", "TEST_5330151", "TEST_5330152"],
    }
    r = session.post(f"{API}/pick-tickets", json=payload)
    assert r.status_code == 200, r.text
    data = r.json()
    assert data["batch_code"] == payload["batch_code"]
    assert data["batch_number"] == 1
    assert data["batch_size"] == 3
    assert data["picker_count"] == 3
    assert data["to_numbers"] == payload["to_numbers"]
    assert "id" in data and isinstance(data["id"], str)
    assert "printed_at" in data

    # GET list should contain it
    r2 = session.get(f"{API}/pick-tickets")
    assert r2.status_code == 200
    items = r2.json()
    assert isinstance(items, list)
    codes = [i["batch_code"] for i in items]
    assert "PTF-Batch-1-TESTX" in codes


def test_check_returns_previous_prints_after_create(session):
    # Ensure a record exists
    payload = {
        "batch_code": "PTF-Batch-2-TESTY",
        "batch_number": 2,
        "batch_size": 3,
        "picker_count": 2,
        "to_numbers": ["TEST_A1", "TEST_A2"],
    }
    r = session.post(f"{API}/pick-tickets", json=payload)
    assert r.status_code == 200

    time.sleep(0.3)
    r2 = session.post(f"{API}/pick-tickets/check", json={"to_numbers": ["TEST_A1", "TEST_A2", "TEST_NOT_EXIST"]})
    assert r2.status_code == 200
    prev = r2.json()["previous_prints"]
    tos = {p["to_number"] for p in prev}
    assert "TEST_A1" in tos
    assert "TEST_A2" in tos
    assert "TEST_NOT_EXIST" not in tos
    for p in prev:
        assert p["batch_code"] == "PTF-Batch-2-TESTY"
        assert isinstance(p["printed_at"], str)
        assert "T" in p["printed_at"]  # ISO format


def test_list_sorted_desc(session):
    # Create two records with small time gap
    p1 = {"batch_code": "PTF-Sort-1", "batch_number": 1, "batch_size": 1, "picker_count": 1, "to_numbers": ["TEST_S1"]}
    p2 = {"batch_code": "PTF-Sort-2", "batch_number": 2, "batch_size": 1, "picker_count": 1, "to_numbers": ["TEST_S2"]}
    session.post(f"{API}/pick-tickets", json=p1)
    time.sleep(0.6)
    session.post(f"{API}/pick-tickets", json=p2)

    r = session.get(f"{API}/pick-tickets")
    assert r.status_code == 200
    items = r.json()
    # find indices
    codes_in_order = [i["batch_code"] for i in items]
    idx1 = codes_in_order.index("PTF-Sort-1")
    idx2 = codes_in_order.index("PTF-Sort-2")
    assert idx2 < idx1, "Newer record should be first (desc order by printed_at)"


def test_double_print_same_tos_flagged(session):
    tos = ["TEST_DBL_1", "TEST_DBL_2"]
    payload1 = {"batch_code": "PTF-Dbl-A", "batch_number": 1, "batch_size": 2, "picker_count": 1, "to_numbers": tos}
    r1 = session.post(f"{API}/pick-tickets", json=payload1)
    assert r1.status_code == 200

    # Check flags them
    rc = session.post(f"{API}/pick-tickets/check", json={"to_numbers": tos})
    assert rc.status_code == 200
    prev = rc.json()["previous_prints"]
    assert len(prev) == 2

    # Print again
    payload2 = {"batch_code": "PTF-Dbl-B", "batch_number": 1, "batch_size": 2, "picker_count": 1, "to_numbers": tos}
    r2 = session.post(f"{API}/pick-tickets", json=payload2)
    assert r2.status_code == 200

    # Now check returns most recent batch_code
    rc2 = session.post(f"{API}/pick-tickets/check", json={"to_numbers": tos})
    prev2 = rc2.json()["previous_prints"]
    assert all(p["batch_code"] == "PTF-Dbl-B" for p in prev2)
