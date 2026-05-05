# Phase 4m Offline Fixtures

This directory is populated locally by `tools/run_offline_tests.py` before the
offline VBA smoke suite runs.

The fixture payload files are intentionally ignored by git. They are copied
from existing `samples/` files or generated as small mock responses, so the
offline tests do not issue HTTP requests.
