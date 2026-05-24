from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

from app.core.settings import RateLimitsCfg
from plugins.jp import edinet_api, edinet_web


def test_edinet_web_selectors_use_spike_confirmed_names() -> None:
    assert edinet_web._SELECTORS == {
        "keyword": 'input[name="W0018vD_KEYWORD"]',
        "kikan": 'select[name="W0018vD_KIKAN"]',
        "syorui1": 'input[name="W0018vCHKSYORUI1"]',
        "search_button": 'input[name="W0018BTNBTN_SEARCH"]',
    }


def test_edinet_web_normalize_row_matches_api_shape() -> None:
    row = edinet_web._normalize_row(
        {
            "SHORUI_KANRI_NO": "S100TR7I",
            "SYORUI_SB_CD_ID": "12001",
            "TEISHUTSU_NICHIJI": "2024/06/25 15:00",
            "SHORUI_NAME": "有価証券報告書（2023/04/01－2024/03/31）",
            "EDINET_CD": "E02144",
            "SHOKEN_CD": "72030",
            "TEISYUTUSYA_NAME": "トヨタ自動車株式会社",
            "PDFKBN": "1",
        },
        ticker_code="7203",
    )

    assert list(row) == list(edinet_api._normalize_row({}))
    assert row == {
        "doc_id": "S100TR7I",
        "doc_type_code": "120",
        "submit_date_time": "2024-06-25 15:00",
        "period_end": "2024-03-31",
        "edinet_code": "E02144",
        "sec_code": "7203",
        "jcn": "",
        "filer_name": "トヨタ自動車株式会社",
        "title": "有価証券報告書（2023/04/01－2024/03/31）",
        "pdf_flag": "1",
    }


def test_edinet_web_extracts_gx_result_rows() -> None:
    payload = {
        "gxValues": [
            {"OTHER": "ignore"},
            {
                "AV125W_RESULT_LIST_JSON": (
                    '[{"SHORUI_KANRI_NO": "S100TR7I"}, {"SHORUI_KANRI_NO": "S100T2RL"}]'
                )
            },
        ]
    }

    rows = edinet_web._extract_result_rows(payload)

    assert [row["SHORUI_KANRI_NO"] for row in rows] == ["S100TR7I", "S100T2RL"]


def test_edinet_web_search_uses_pdf_renderer_pool(monkeypatch) -> None:
    calls: dict[str, object] = {}
    payload = {
        "gxValues": [
            {
                "AV125W_RESULT_LIST_JSON": (
                    """
                    [
                      {
                        "SHORUI_KANRI_NO": "S100TR7I",
                        "SYORUI_SB_CD_ID": "12001",
                        "TEISHUTSU_NICHIJI": "2024/06/25 15:00",
                        "SHORUI_NAME": "有価証券報告書（2023/04/01－2024/03/31）",
                        "EDINET_CD": "E02144",
                        "PDFKBN": "1"
                      },
                      {
                        "SHORUI_KANRI_NO": "S100OLD",
                        "SYORUI_SB_CD_ID": "12001",
                        "TEISHUTSU_NICHIJI": "2023/06/25 15:00"
                      }
                    ]
                    """
                )
            }
        ]
    }

    class _Limiter:
        def acquire_blocking(self) -> None:
            calls["rate_acquired"] = True

    class _Request:
        method = "POST"
        post_data = ""

    class _Response:
        url = "https://disclosure2.edinet-fsa.go.jp/WEEE0030.aspx?GXAjaxRequest=1"
        status = 200
        headers = {"content-type": "application/json"}
        request = _Request()

        def json(self):
            return payload

    class _Page:
        def __init__(self) -> None:
            self.handler = None
            self.closed = False

        def on(self, event: str, handler) -> None:
            assert event == "response"
            self.handler = handler

        def goto(self, url: str, **_kwargs) -> None:
            calls["goto"] = url

        def fill(self, selector: str, value: str) -> None:
            calls["fill"] = (selector, value)

        def select_option(self, selector: str, value: str) -> None:
            calls["select"] = (selector, value)

        def is_checked(self, selector: str) -> bool:
            calls["is_checked"] = selector
            return False

        def check(self, selector: str) -> None:
            calls["check"] = selector

        def click(self, selector: str) -> None:
            calls["click"] = selector
            assert self.handler is not None
            self.handler(_Response())

        def wait_for_timeout(self, _timeout_ms: int) -> None:
            calls["waited"] = True

        def wait_for_function(self, *_args, **_kwargs) -> None:
            calls["wait_for_function"] = True

        def close(self) -> None:
            self.closed = True
            calls["page_closed"] = True

    class _Context:
        def __init__(self) -> None:
            self.page = _Page()

        def new_page(self):
            return self.page

        def clear_cookies(self) -> None:
            calls["cookies_cleared"] = True

        def clear_permissions(self) -> None:
            calls["permissions_cleared"] = True

        def close(self) -> None:
            calls["context_closed"] = True

    class _Browser:
        def __init__(self) -> None:
            self.context = _Context()

        def new_context(self, **kwargs):
            calls["context_kwargs"] = kwargs
            return self.context

    state = SimpleNamespace(browser=_Browser(), context=None, context_user_agent=None)

    from app.core import pdf_renderer

    monkeypatch.setattr(edinet_web, "limiter", lambda source, rate: _Limiter())
    monkeypatch.setattr(edinet_web, "_edinet_web_rate", lambda: 99.0)
    monkeypatch.setattr(pdf_renderer, "_ensure_state", lambda: state)
    monkeypatch.setattr(pdf_renderer, "_close_context", lambda _state: None)
    monkeypatch.setattr(pdf_renderer, "_clear_context_state", lambda _context: True)

    rows = edinet_web.search_filings("7203", 2024)

    assert [row["doc_id"] for row in rows] == ["S100TR7I"]
    assert calls["fill"] == (edinet_web._SELECTORS["keyword"], "7203")
    assert calls["select"] == (edinet_web._SELECTORS["kikan"], "7")
    assert calls["check"] == edinet_web._SELECTORS["syorui1"]
    assert calls["click"] == edinet_web._SELECTORS["search_button"]
    assert calls["context_kwargs"]["locale"] == "ja-JP"
    assert calls["rate_acquired"] is True
    assert calls["page_closed"] is True


def test_edinet_web_download_document_pdf_uses_direct_url(monkeypatch, tmp_path: Path) -> None:
    calls: dict[str, object] = {}

    class _Client:
        def __enter__(self):
            calls["client_entered"] = True
            return self

        def __exit__(self, *_args):
            return False

    def fake_stream_to_file(_client, url, dest, **kwargs):
        calls["url"] = url
        calls["kwargs"] = kwargs
        dest.write_bytes(b"%PDF\nbody")
        return 9

    def fake_default_client(**kwargs):
        calls["client_kwargs"] = kwargs
        return _Client()

    monkeypatch.setattr(edinet_web, "default_client", fake_default_client)
    monkeypatch.setattr(edinet_web, "stream_to_file", fake_stream_to_file)
    monkeypatch.setattr(edinet_web, "_edinet_web_rate", lambda: 99.0)

    n_bytes = edinet_web.download_document_pdf("S100TR7I", tmp_path / "report.pdf")

    assert n_bytes == 9
    assert calls["client_kwargs"] == {"source": "edinet_web", "timeout": 120.0}
    assert calls["url"] == "https://disclosure2dl.edinet-fsa.go.jp/searchdocument/pdf/S100TR7I.pdf"
    assert calls["kwargs"]["source"] == "edinet_web"
    assert calls["kwargs"]["rate"] == 99.0
    assert calls["kwargs"]["read_timeout"] == 180.0


def test_edinet_web_rate_limit_default_is_one_qps() -> None:
    assert RateLimitsCfg().edinet_web == 1.0
