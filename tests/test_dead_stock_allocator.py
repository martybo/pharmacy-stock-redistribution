import pytest

pd = pytest.importorskip("pandas")

from dead_stock_allocator import allocate_for_product


def test_same_am_receivers_prioritized_before_cross_am():
    row = pd.Series(
        {
            "st::Source": 10,
            "usage::SameAM": 5,
            "st::SameAM": 0,
            "usage::CrossAM": 5,
            "st::CrossAM": 0,
        }
    )
    store_to_am = {
        "Source": "AM-1",
        "SameAM": "AM-1",
        "CrossAM": "AM-2",
    }

    allocations = allocate_for_product(
        row,
        source_store="Source",
        store_to_am=store_to_am,
        stores=["Source", "SameAM", "CrossAM"],
        am_override={},
    )

    same_am_allocs = [a for a in allocations if a[0] == "SameAM"]
    cross_am_allocs = [a for a in allocations if a[0] == "CrossAM"]

    assert same_am_allocs, "Expected same-AM receiver to obtain allocation"
    assert cross_am_allocs, "Expected cross-AM receiver to obtain allocation"

    assert same_am_allocs[0][1] == 5
    assert cross_am_allocs[0][1] == 5

    first_cross_index = min(i for i, a in enumerate(allocations) if a[0] == "CrossAM")
    first_same_index = min(i for i, a in enumerate(allocations) if a[0] == "SameAM")
    assert first_same_index < first_cross_index
