from __future__ import annotations

import datetime as dt

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from app.core.models import Period, PeriodType
from app.ui import strings as ui_strings
from app.ui.i18n import LanguageManager


class PeriodSelector(QFrame):
    """Year range + period type checkboxes."""

    selection_changed = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("Card")
        self.setProperty("class", "Card")
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        header = QWidget()
        header.setObjectName("CardHeader")
        h = QHBoxLayout(header)
        h.setContentsMargins(20, 16, 20, 8)
        h.setSpacing(12)

        self.title_label = QLabel(ui_strings.PS_TITLE)
        self.title_label.setObjectName("CardTitle")
        self.subtitle_label = QLabel(ui_strings.PS_SUBTITLE)
        self.subtitle_label.setObjectName("CardSubtitle")
        h.addWidget(self.title_label)
        h.addWidget(self.subtitle_label)
        h.addStretch(1)
        outer.addWidget(header)

        body = QWidget()
        body.setObjectName("CardBody")
        b = QVBoxLayout(body)
        b.setContentsMargins(20, 0, 20, 18)
        b.setSpacing(14)

        # Year range
        year_row = QHBoxLayout()
        year_row.setSpacing(10)
        self.from_year_label = QLabel(ui_strings.PS_FROM_YEAR)
        year_row.addWidget(self.from_year_label)
        self.from_year = QComboBox()
        self.to_year = QComboBox()
        cur = dt.date.today().year
        for y in range(cur - 10, cur + 1):
            self.from_year.addItem(str(y), y)
            self.to_year.addItem(str(y), y)
        # Default: last 3 fiscal years (e.g. 2022..2024)
        self.from_year.setCurrentText(str(cur - 3))
        self.to_year.setCurrentText(str(cur - 1))
        self.from_year.currentIndexChanged.connect(lambda *_: self.selection_changed.emit())
        self.to_year.currentIndexChanged.connect(lambda *_: self.selection_changed.emit())
        self.from_year.setMinimumWidth(110)
        self.to_year.setMinimumWidth(110)
        year_row.addWidget(self.from_year)
        year_row.addSpacing(8)
        dash = QLabel("→")
        dash.setObjectName("MutedArrow")
        year_row.addWidget(dash)
        year_row.addSpacing(8)
        self.to_year_label = QLabel(ui_strings.PS_TO_YEAR)
        year_row.addWidget(self.to_year_label)
        year_row.addWidget(self.to_year)
        year_row.addStretch(1)
        b.addLayout(year_row)

        # Period checkboxes
        type_row = QHBoxLayout()
        type_row.setSpacing(20)
        self.type_label = QLabel(ui_strings.PS_TYPE_LABEL)
        self.type_label.setObjectName("ConfigLabel")
        type_row.addWidget(self.type_label)

        self.cb_annual = QCheckBox(ui_strings.PS_ANNUAL)
        self.cb_annual.setChecked(True)
        self.cb_q3 = QCheckBox(ui_strings.PS_Q3)
        self.cb_q2 = QCheckBox(ui_strings.PS_Q2)
        self.cb_q1 = QCheckBox(ui_strings.PS_Q1)
        self.cb_ipo = QCheckBox(ui_strings.PS_IPO)
        for cb in (self.cb_annual, self.cb_q3, self.cb_q2, self.cb_q1, self.cb_ipo):
            cb.toggled.connect(lambda *_: self.selection_changed.emit())
            type_row.addWidget(cb)
        type_row.addStretch(1)
        b.addLayout(type_row)

        outer.addWidget(body)
        LanguageManager.instance().language_changed.connect(self._retranslate)

    def periods(self) -> list[Period]:
        from_y = int(self.from_year.currentData())
        to_y = int(self.to_year.currentData())
        if from_y > to_y:
            from_y, to_y = to_y, from_y
        types: list[PeriodType] = []
        if self.cb_annual.isChecked():
            types.append(PeriodType.ANNUAL)
        if self.cb_q1.isChecked():
            types.append(PeriodType.Q1)
        if self.cb_q2.isChecked():
            types.append(PeriodType.Q2)
        if self.cb_q3.isChecked():
            types.append(PeriodType.Q3)
        if self.cb_ipo.isChecked():
            types.append(PeriodType.IPO_PROSPECTUS)
        out: list[Period] = []
        for y in range(from_y, to_y + 1):
            for t in types:
                out.append(Period(year=y, type=t))
        return out

    def _retranslate(self, _lang: str = "") -> None:
        self.title_label.setText(ui_strings.PS_TITLE)
        self.subtitle_label.setText(ui_strings.PS_SUBTITLE)
        self.from_year_label.setText(ui_strings.PS_FROM_YEAR)
        self.to_year_label.setText(ui_strings.PS_TO_YEAR)
        self.type_label.setText(ui_strings.PS_TYPE_LABEL)
        self.cb_annual.setText(ui_strings.PS_ANNUAL)
        self.cb_q3.setText(ui_strings.PS_Q3)
        self.cb_q2.setText(ui_strings.PS_Q2)
        self.cb_q1.setText(ui_strings.PS_Q1)
        self.cb_ipo.setText(ui_strings.PS_IPO)
