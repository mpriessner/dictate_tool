#!/usr/bin/env python3
# focus_dashboard.py
#
# Companion dashboard for FocusTimer
# ----------------------------------
# • Shows a donut chart (Focus vs Break) for the chosen Day / Week / Month
# • Lists the Top‑10 categories with proportional horizontal bars
#
# Author: 2025‑04‑22

import os, csv, sys
from datetime import datetime
from pathlib import Path
from collections import defaultdict

from PyQt5.QtCore   import Qt, QDate, QRectF, QPointF
from PyQt5.QtGui    import QColor, QPainter, QPen, QBrush, QFont, QPainterPath
from PyQt5.QtWidgets import (
    QApplication, QDialog, QLabel, QWidget, QPushButton, QVBoxLayout, QHBoxLayout,
    QFrame, QDateEdit, QScrollArea
)


# ──────────────────────────────────────────────────────────────────────────────
#  Donut chart
# ──────────────────────────────────────────────────────────────────────────────
class DonutChartWidget(QWidget):
    """Simple interactive donut (pie‑with‑hole) chart."""

    COLORS = {
        'Focus': QColor("#89B4FA"),
        'Break': QColor("#74C7EC"),
    }

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(200, 200)
        self.data: dict[str, int] = {}
        self.hover = None
        self.paths: dict[str, QPainterPath] = {}
        self.setMouseTracking(True)

    # ── public ────────────────────────────────────────────────────────────
    def setData(self, data: dict[str, int]):
        self.data  = {k: v for k, v in data.items() if v > 0}
        self.hover = None
        self.paths.clear()
        self.setToolTip("")
        self.update()

    # ── drawing ───────────────────────────────────────────────────────────
    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)

        if not self.data:
            p.setPen(QColor("#777"))
            p.setFont(QFont("Arial", 10))
            p.drawText(self.rect(), Qt.AlignCenter, "No Data")
            return

        r   = self.rect().adjusted(10, 10, -10, -10)
        cen = r.center()
        R   = min(r.width(), r.height()) / 2
        r0  = R * 0.60                         # inner radius

        total = sum(self.data.values())
        start_deg16 = 90 * 16                  # 12 o'clock

        self.paths.clear()
        for label, val in self.data.items():
            span_deg16 = val / total * 360 * 16
            if span_deg16 <= 0:                 # skip zero slices
                continue

            # outer slice
            slice_rect = QRectF(cen.x()-R, cen.y()-R, 2*R, 2*R)
            path = QPainterPath()
            path.moveTo(cen)
            path.arcTo(slice_rect, start_deg16 / 16, span_deg16 / 16)
            path.closeSubpath()

            # cut inner hole
            hole = QPainterPath()
            hole.addEllipse(QRectF(cen.x()-r0, cen.y()-r0, 2*r0, 2*r0))
            slice_path = path.subtracted(hole)

            self.paths[label] = slice_path

            # colour & hover
            color = self.COLORS.get(label, QColor("gray"))
            if label == self.hover:
                color = color.lighter(120)
                p.setPen(QPen(Qt.white, 1))
            else:
                p.setPen(Qt.NoPen)
            p.setBrush(QBrush(color))
            p.drawPath(slice_path)

            start_deg16 += span_deg16

        # draw inner hole (background‑coloured disk)
        p.setPen(Qt.NoPen)
        p.setBrush(QColor("#1E1E2E"))
        p.drawEllipse(QPointF(*cen.toTuple()), r0, r0)

    # ── interaction ───────────────────────────────────────────────────────
    def mouseMoveEvent(self, evt):
        pos = evt.pos()
        hovered = None
        for label, path in self.paths.items():
            if path.contains(pos):
                hovered = label
                break
        if hovered != self.hover:
            self.hover = hovered
            if hovered:
                s = self.data[hovered]
                h, m = divmod(int(s//60), 60)
                self.setToolTip(f"{hovered}: {h} h {m} m")
            else:
                self.setToolTip("")
            self.update()

    def leaveEvent(self, _):
        self.hover = None
        self.setToolTip("")
        self.update()


# ──────────────────────────────────────────────────────────────────────────────
#  Weekly Bar Chart Widget
# ──────────────────────────────────────────────────────────────────────────────
class WeeklyBarChartWidget(QWidget):
    """A widget to display daily work hours as a bar chart for the week view."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(400, 200)
        self.daily_hours = {}  # Format: {'Mon': hours, 'Tue': hours, ...}
        self.max_hours = 8  # Default max hours (y-axis scale)
        self.day_labels = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
        self.colors = {
            'Mon': QColor("#89B4FA"),
            'Tue': QColor("#89B4FA"),
            'Wed': QColor("#89B4FA"),
            'Thu': QColor("#89B4FA"),
            'Fri': QColor("#89B4FA"),
            'Sat': QColor("#89B4FA"),
            'Sun': QColor("#89B4FA")
        }
        self.hover_color = QColor("#B4BEFE")
        self.hovered_bar = None
        self.setMouseTracking(True)
        
    def update_data(self, daily_hours, max_hours=None):
        """Update the chart data and redraw."""
        self.daily_hours = daily_hours
        if max_hours is not None:
            self.max_hours = max_hours
        else:
            # Auto-calculate max_hours based on data (with some padding)
            if daily_hours:
                highest_value = max(daily_hours.values()) if daily_hours else 0
                self.max_hours = max(8, round(highest_value * 1.2))  # At least 8 hours, with 20% padding
        self.update()
        
    def format_time(self, hours):
        """Format hours as 'X hr Y min'."""
        hours_int = int(hours)
        minutes = int((hours - hours_int) * 60)
        return f"{hours_int} hr {minutes} min"
        
    def paintEvent(self, event):
        """Draw the bar chart."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Get widget dimensions
        width = self.width()
        height = self.height()
        
        # Calculate chart area
        chart_margin = 40  # Margin for labels
        chart_width = width - chart_margin
        chart_height = height - chart_margin - 20  # Extra space at bottom for day labels
        
        # Draw y-axis and horizontal grid lines
        painter.setPen(QPen(QColor("#6C7086"), 1))
        
        # Draw y-axis labels and grid lines
        for i in range(5):  # 0, 2, 4, 6, 8 hours (or scaled based on max_hours)
            y_value = i * (self.max_hours / 4)
            y_pos = height - chart_margin - (y_value / self.max_hours) * chart_height
            
            # Draw grid line - convert float y_pos to int
            painter.drawLine(int(chart_margin), int(y_pos), int(width - 10), int(y_pos))
            
            # Draw y-axis label - convert float y_pos to int
            painter.drawText(QPointF(5, int(y_pos + 5)), f"{y_value:.0f}")
        
        # Draw the average line
        if self.daily_hours:
            avg_hours = sum(self.daily_hours.values()) / len(self.daily_hours)
            avg_y = height - chart_margin - (avg_hours / self.max_hours) * chart_height
            
            # Draw dashed average line - convert float avg_y to int
            dash_pen = QPen(QColor("#F5C2E7"), 1, Qt.DashLine)
            painter.setPen(dash_pen)
            painter.drawLine(int(chart_margin), int(avg_y), int(width - 10), int(avg_y))
            
            # Draw "Average" label - convert float avg_y to int
            painter.drawText(QPointF(int(width - 70), int(avg_y - 5)), "Average")
        
        # Draw bars for each day
        if self.daily_hours:
            bar_width = (chart_width - chart_margin) / 7  # 7 days in a week
            bar_spacing = bar_width * 0.2
            bar_width = bar_width * 0.8
            
            for i, day in enumerate(self.day_labels):
                hours = self.daily_hours.get(day, 0)
                
                # Calculate bar position
                bar_x = chart_margin + i * (bar_width + bar_spacing)
                bar_height = (hours / self.max_hours) * chart_height
                bar_y = height - chart_margin - bar_height
                
                # Draw the bar - convert float coordinates to int
                bar_color = self.hover_color if day == self.hovered_bar else self.colors[day]
                painter.setBrush(QBrush(bar_color))
                painter.setPen(Qt.NoPen)
                painter.drawRect(int(bar_x), int(bar_y), int(bar_width), int(bar_height))
                
                # Draw day label - convert float coordinates to int
                painter.setPen(QPen(QColor("#CDD6F4")))
                painter.drawText(QPointF(int(bar_x + bar_width/2 - 10), int(height - 10)), day)
                
                # Draw hours on top of the bar if there's data - convert float coordinates to int
                if hours > 0:
                    painter.drawText(QPointF(int(bar_x + bar_width/2 - 15), int(bar_y - 5)), f"{hours:.1f}")
    
    def mouseMoveEvent(self, event):
        """Handle mouse hover to highlight bars."""
        if not self.daily_hours:
            return
            
        # Calculate bar dimensions
        width = self.width()
        height = self.height()
        chart_margin = 40
        chart_width = width - chart_margin
        bar_width = (chart_width - chart_margin) / 7
        bar_spacing = bar_width * 0.2
        bar_width = bar_width * 0.8
        
        # Check which bar is being hovered
        mouse_x = event.x()
        old_hovered = self.hovered_bar
        self.hovered_bar = None
        
        for i, day in enumerate(self.day_labels):
            bar_x = chart_margin + i * (bar_width + bar_spacing)
            if bar_x <= mouse_x <= bar_x + bar_width:
                self.hovered_bar = day
                hours = self.daily_hours.get(day, 0)
                self.setToolTip(f"{day}: {self.format_time(hours)}")
                break
        
        # Only update if the hover state changed
        if old_hovered != self.hovered_bar:
            self.update()
            
    def leaveEvent(self, event):
        """Reset hover state when mouse leaves the widget."""
        if self.hovered_bar:
            self.hovered_bar = None
            self.update()


# ──────────────────────────────────────────────────────────────────────────────
#  Horizontal bar per category
# ──────────────────────────────────────────────────────────────────────────────
class CategoryBarWidget(QWidget):
    def __init__(self, label: str, secs: int, max_secs: int, color: QColor, parent=None):
        super().__init__(parent)
        self.label, self.secs, self.max = label, secs, max_secs
        self.color = color
        self.setFixedHeight(26)
        self.setToolTip(self.fmt(secs))

    def fmt(self, s):
        h, m = divmod(int(s//60), 60)
        return f"{h}h {m}m" if h else f"{m}m"

    # draw
    def paintEvent(self, _):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        w, h = self.width(), self.height()
        bw   = w * 0.70 * self.secs / self.max if self.max else 0
        bar  = QRectF(0, 4, bw, h‑8)
        p.setPen(Qt.NoPen); p.setBrush(self.color)
        p.drawRoundedRect(bar, 4, 4)

        # label
        p.setPen(QColor("#CDD6F4"))
        p.setFont(QFont("Arial", 9))
        txt = f"{self.label[:15]}{'…' if len(self.label) > 15 else ''} ({self.fmt(self.secs)})"
        p.drawText(QRectF(w*0.73, 0, w*0.27, h), Qt.AlignLeft | Qt.AlignVCenter, txt)


# ──────────────────────────────────────────────────────────────────────────────
#  Dashboard main window
# ──────────────────────────────────────────────────────────────────────────────
class DashboardWindow(QDialog):
    """Day / Week / Month productivity dashboard."""

    def __init__(self, parent=None, log_dir: str | Path = None):
        super().__init__(parent)
        self.log_dir  = Path(log_dir or ".").expanduser()
        self.viewMode = "day"
        self.curDate  = QDate.currentDate()

        self.setWindowTitle("FocusTimer Dashboard")
        self.setMinimumSize(820, 640)
        self.setStyleSheet("""
            QDialog { background:#1E1E2E; color:#CDD6F4; }
            QLabel  { color:#CDD6F4; }
            QFrame#card { background:#28283D; border-radius:8px; }
            QPushButton { background:#313244; color:#CDD6F4; border:none;
                          padding:5px 10px; border-radius:4px; min-width:60px; }
            QPushButton:hover { background:#45475A; }
            QPushButton:checked { background:#89B4FA; color:#1E1E2E; }
            QDateEdit { background:#313244; color:#CDD6F4; border:1px solid #555;
                        padding:3px; border-radius:3px; min-width:90px; }
        """)

        self.buildUI()
        self.load()                         # initial populate

    # ── UI ────────────────────────────────────────────────────────────────
    def buildUI(self):
        root = QVBoxLayout(self); root.setContentsMargins(20,20,20,20); root.setSpacing(18)

        # header: title + view buttons + date nav
        hdr = QHBoxLayout(); root.addLayout(hdr)
        title = QLabel("Productivity Dashboard"); title.setFont(QFont("Arial", 18, QFont.Bold))
        hdr.addWidget(title); hdr.addStretch()

        self.dayBtn   = self.makeViewBtn("Day")
        self.weekBtn  = self.makeViewBtn("Week")
        self.monthBtn = self.makeViewBtn("Month", checked=False)
        for b in (self.dayBtn, self.weekBtn, self.monthBtn): hdr.addWidget(b)

        hdr.addSpacing(15)
        self.prevBtn = QPushButton("←"); self.prevBtn.setFixedWidth(28)
        self.nextBtn = QPushButton("→"); self.nextBtn.setFixedWidth(28)
        self.dateEdit= QDateEdit(self.curDate); self.dateEdit.setCalendarPopup(True)
        hdr.addWidget(self.prevBtn); hdr.addWidget(self.dateEdit); hdr.addWidget(self.nextBtn)

        self.prevBtn.clicked.connect(lambda: self.shiftDate(-1))
        self.nextBtn.clicked.connect(lambda: self.shiftDate(+1))
        self.dateEdit.dateChanged.connect(self.onDateChanged)

        # --- top row -----------------------------------------------------------------
        top = QHBoxLayout(); top.setSpacing(15); root.addLayout(top)

        # summary card
        card = QFrame(); card.setObjectName("card"); card.setMinimumWidth(200)
        tlay = QVBoxLayout(card); tlay.setContentsMargins(15,15,15,15)
        lbl = QLabel("Daily Summary"); lbl.setFont(QFont("Arial", 14, QFont.Bold))
        self.workLbl  = QLabel(); self.focusLbl = QLabel()
        for w in (lbl, self.workLbl, self.focusLbl): tlay.addWidget(w)
        tlay.addStretch()
        top.addWidget(card, 1)

        # charts (donut for day/month, bar chart for week)
        chartCard = QFrame(); chartCard.setObjectName("card")
        clay = QVBoxLayout(chartCard)
        self.chartTitle = QLabel("Time Breakdown"); self.chartTitle.setAlignment(Qt.AlignCenter)
        self.chartTitle.setFont(QFont("Arial", 14, QFont.Bold))
        self.donut = DonutChartWidget()
        self.weekly_chart = WeeklyBarChartWidget()
        clay.addWidget(self.chartTitle)
        clay.addWidget(self.donut)
        clay.addWidget(self.weekly_chart)
        self.weekly_chart.hide()  # Initially hide weekly chart
        top.addWidget(chartCard, 2)

        # --- bottom: categories -------------------------------------------------------
        catCard = QFrame(); catCard.setObjectName("card")
        blay = QVBoxLayout(catCard); blay.setContentsMargins(15,15,15,15)
        blay.addWidget(QLabel("Top Categories", font=QFont("Arial", 14, QFont.Bold)))
        self.catBox = QVBoxLayout(); self.catBox.setSpacing(4)
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        cont   = QWidget(); cont.setLayout(self.catBox); scroll.setWidget(cont)
        blay.addWidget(scroll); blay.addStretch()
        root.addWidget(catCard)

    def makeViewBtn(self, text, checked=True):
        b = QPushButton(text); b.setCheckable(True); b.setChecked(checked)
        b.clicked.connect(lambda _=False, m=text.lower(): self.switchMode(m))
        return b

    # ── date helpers ───────────────────────────────────────────────────────
    def shiftDate(self, delta):
        d = self.dateEdit.date()
        if   self.viewMode == "day":   d = d.addDays(delta)
        elif self.viewMode == "week":  d = d.addDays(7*delta)
        else:                          d = d.addMonths(delta)
        if d <= QDate.currentDate():   self.dateEdit.setDate(d)

    def onDateChanged(self, d): self.curDate = d; self.load()

    def switchMode(self, m):
        if m == self.viewMode: return
        self.viewMode = m
        for b, name in ((self.dayBtn,"day"), (self.weekBtn,"week"), (self.monthBtn,"month")):
            b.setChecked(name == m)
        fmt = {"day":"yyyy‑MM‑dd", "week":"yyyy 'W'ww", "month":"MMM yyyy"}[m]
        self.dateEdit.setDisplayFormat(fmt)
        
        # Show/hide appropriate chart based on view mode
        if m == "week":
            self.chartTitle.setText("Work Hours by Day")
            self.donut.hide()
            self.weekly_chart.show()
        else:
            self.chartTitle.setText("Time Breakdown" if m == "day" else "Monthly Time Breakdown")
            self.donut.show()
            self.weekly_chart.hide()
            
        self.load()

    # ── loader ----------------------------------------------------------------
    def load(self):
        self.clearUI()

        if self.viewMode == "day":
            dates = [self.curDate]
        elif self.viewMode == "week":
            mon = self.curDate.addDays(1-self.curDate.dayOfWeek())
            dates= [mon.addDays(i) for i in range(7) if mon.addDays(i)<=QDate.currentDate()]
        else:                             # month
            days = self.curDate.daysInMonth()
            dates= [QDate(self.curDate.year(), self.curDate.month(), d)
                    for d in range(1, days+1)
                    if QDate(self.curDate.year(), self.curDate.month(), d)<=QDate.currentDate()]

        totHrs, focusS, breakS = 0.0, 0, 0
        apps = defaultdict(int)
        
        # For weekly bar chart
        daily_hours = {}

        for d in dates:
            f = self.log_dir / d.toString("yyyy-MM-dd") / "activity_log.csv"
            if not f.exists(): continue

            try:
                with f.open() as csvfile:
                    rdr = list(csv.reader(csvfile)); 
                    if len(rdr)<=1: continue          # only header

                    # last row col‑3 = cumulative work_hours
                    try: 
                        day_hours = float(rdr[-1][3])
                        totHrs += day_hours
                        
                        # Store daily hours for the weekly bar chart
                        if self.viewMode == "week":
                            day_name = d.toString("ddd")[:3]  # Get short day name (Mon, Tue, etc.)
                            daily_hours[day_name] = day_hours
                    except: pass

                    prev_ts = None
                    for row in rdr[1:]:
                        if len(row)<5: continue
                        ts  = datetime.strptime(row[0], "%Y-%m-%d %H:%M:%S")
                        act = int(row[1]) if row[1].isdigit() else 0
                        app = row[4] or "Unknown"

                        dur = 0
                        if prev_ts: dur = (ts-prev_ts).total_seconds()
                        prev_ts = ts
                        if dur<=0: continue

                        if act:
                            focusS += dur;    apps[app]+=dur
                        else:
                            breakS += dur;    apps["Break"]+=dur
            except Exception as e:
                print("Error reading", f, e)

        # summary labels ----------------------------------------------------
        h, m = divmod(int(totHrs*60), 60)
        self.workLbl.setText(f"Work Hours: {h} h {m} m")
        totalS = focusS+breakS
        pct = (focusS/totalS*100) if totalS else 0
        self.focusLbl.setText(f"Focus: {pct:.0f}%")

        self.donut.setData({"Focus":focusS, "Break":breakS})
        
        # Update weekly bar chart if in week view
        if self.viewMode == "week":
            self.weekly_chart.update_data(daily_hours)

        # categories --------------------------------------------------------
        if apps:
            sorted_apps = sorted(apps.items(), key=lambda x: x[1], reverse=True)[:10]
            max_s = sorted_apps[0][1]
            palette = ["#F2CDCD","#DDB6F2","#F5C2E7","#E8A2AF","#F28FAD",
                       "#ABE9B3","#FAE3B0","#F8BD96","#EE99A0","#89DCEB"]
            for i,(name, secs) in enumerate(sorted_apps):
                self.catBox.addWidget(CategoryBarWidget(name, secs, max_s,
                                     QColor(palette[i%len(palette)])))

        if not apps:
            self.catBox.addWidget(QLabel("No data available."))

    def clearUI(self):
        self.workLbl.setText("Loading…"); self.focusLbl.clear()
        self.donut.setData({})
        self.weekly_chart.update_data({})
        while self.catBox.count():
            w = self.catBox.takeAt(0).widget()
            if w: w.deleteLater()


# ──────────────────────────────────────────────────────────────────────────────
#  Dev / manual run
# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    logbase = Path(sys.argv[1] if len(sys.argv)>1 else "./focus_logs")
    app = QApplication(sys.argv)
    dlg = DashboardWindow(log_dir=logbase)
    dlg.show()
    sys.exit(app.exec_())
