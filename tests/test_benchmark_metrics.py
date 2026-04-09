from outlook_mail_assistant.benchmarking import measure_peak_memory_mb


def test_measure_peak_memory_mb_returns_numeric_value():
    value = measure_peak_memory_mb(lambda: [str(i) for i in range(1000)])

    assert isinstance(value, float)
    assert value >= 0.0
