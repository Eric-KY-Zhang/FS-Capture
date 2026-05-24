from __future__ import annotations


def pytest_addoption(parser):
    parser.addoption(
        "--benchmark-only",
        action="store_true",
        default=False,
        help="Run only tests marked with @pytest.mark.benchmark.",
    )


def pytest_collection_modifyitems(config, items):
    if not config.getoption("--benchmark-only"):
        return

    selected = []
    deselected = []
    for item in items:
        if item.get_closest_marker("benchmark") is not None:
            selected.append(item)
        else:
            deselected.append(item)

    if deselected:
        config.hook.pytest_deselected(items=deselected)
        items[:] = selected
