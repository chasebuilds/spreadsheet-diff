from benchmark.generate import generate
from benchmark.run_benchmark import score


def test_generator_is_seeded() -> None:
    assert generate(1, 50).new.rows == generate(1, 50).new.rows
    assert generate(1, 50).new.rows != generate(2, 50).new.rows


def test_small_benchmark_run() -> None:
    s = score(3, 400)
    for t in (s.added, s.removed, s.cells, s.changed_rows):
        assert t[0] == t[1] == t[2]
    assert s.noise_flagged == 0 and s.noise_flagged_strict == s.noise
    assert s.columns_ok
