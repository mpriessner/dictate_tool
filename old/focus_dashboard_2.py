#!/usr/bin/env python3
#
# focus_dashboard.py  ⟵ updated with improved visualizations
#
# NEW: Enhanced visualization in Day/Week/Month views
# ────────────────────────────────────────────────────────────────────────────
#  Day  ➜  donut  +  horizontal timeline showing every Focus / Break block
#  Week ➜  bar chart showing daily work hours
#  Month  ➜  weekly breakdown showing hours per week
#
#  *Requires only PyQt5 – no matplotlib – so nothing extra to install.*

import os, csv, sys
from datetime import datetime, timedelta
from pathlib import Path
from collections import defaultdict

from PyQt5.QtCore import Qt, QDate, QRectF, QPointF, QSize
from PyQt5.QtGui import QColor, QPainter, QPen, QBrush, QFont, QPainterPath
from PyQt5.QtWidgets import (
    QApplication, QDialog, QLabel, QWidget, QPushButton, QVBoxLayout,
    QHBoxLayout, QFrame, QDateEdit, QScrollArea
)

#─────────────────────────  COLOUR CONSTANTS  ────────────────────────────────
CLR_FOCUS = QColor("#89B4FA")
CLR_BREAK = QColor("#74C7EC")
CLR_BG    = QColor("#1E1E2E")


#─────────────────────────  DONUT CHART  ─────────────────────────────────────
class DonutChartWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.data, self.paths, self.hover = {}, {}, None
        self.setMouseTracking(True)
        self.setMinimumSize(180, 180)

    def setData(self, data: dict[str, int]):
        self.data = {k: v for k, v in data.items() if v > 0}
        self.paths.clear(); self.hover = None; self.setToolTip(""); self.update()

    def paintEvent(self, _):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        if not self.data:
            p.setPen(QColor("#666")); p.setFont(QFont("Arial", 10))
            p.drawText(self.rect(), Qt.AlignCenter, "No Data"); return

        r = self.rect().adjusted(12, 12, -12, -12)
        cen, R, r0 = r.center(), min(r.width(), r.height())/2, min(r.width(), r.height())/2*0.6
        start = 90*16; total = sum(self.data.values()); self.paths.clear()
        for lbl,val in self.data.items():
            span = val/total*360*16
            rect = QRectF(cen.x()-R,cen.y()-R,2*R,2*R)
            path = QPainterPath(); path.moveTo(cen); path.arcTo(rect,start/16,span/16)
            hole = QPainterPath(); hole.addEllipse(QRectF(cen.x()-r0,cen.y()-r0,2*r0,2*r0))
            seg  = path.subtracted(hole); self.paths[lbl]=seg
            col  = CLR_FOCUS if lbl=="Focus" else CLR_BREAK
            if lbl==self.hover: col=col.lighter(120); p.setPen(QPen(Qt.white,1))
            else:               p.setPen(Qt.NoPen)
            p.setBrush(QBrush(col)); p.drawPath(seg); start+=span
        # draw centre hole
        p.setPen(Qt.NoPen); p.setBrush(QBrush(CLR_BG))
        p.drawEllipse(QPointF(cen), r0, r0)

    def mouseMoveEvent(self, e):
        pos,label = e.pos(),None
        for k,v in self.paths.items():
            if v.contains(pos): label=k; break
        if label!=self.hover:
            self.hover=label
            if label:
                s = self.data[label]; h,m = divmod(int(s//60),60)
                self.setToolTip(f"{label}: {h} h {m} m")
            else: self.setToolTip("")
            self.update()

    def leaveEvent(self,_): self.hover=None; self.setToolTip(""); self.update()


#─────────────────────────  TIMELINE BAR  ────────────────────────────────────
class TimelineWidget(QWidget):
    """24‑hour horizontal bar of coloured blocks (Focus / Break)."""
    def __init__(self,parent=None): 
        super().__init__(parent); 
        self.intervals=[]
    
    def setData(self, intervals:list[tuple[datetime,datetime,bool]]):
        self.intervals=intervals; 
        self.update()
    
    def sizeHint(self): 
        return self.minimumSizeHint()
    
    def minimumSizeHint(self): 
        return QSize(300,60)

    def paintEvent(self,_):
        p=QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        W,H = self.width(), self.height()
        barH, y = H*0.6, (H-H*0.6)/2
        # background
        p.fillRect(0,y,W,barH,QColor("#272738"))
        if not self.intervals: return
        secs_in_day=24*3600
        for st,en,focus in self.intervals:
            x0= (st.hour*3600+st.minute*60+st.second)/secs_in_day*W
            x1= (en.hour*3600+en.minute*60+en.second)/secs_in_day*W
            col= CLR_FOCUS if focus else CLR_BREAK
            p.setPen(Qt.NoPen); p.setBrush(col)
            p.drawRect(QRectF(x0,y, max(1,x1-x0), barH))
        # hour ticks
        p.setPen(QColor("#444")); p.setBrush(Qt.NoBrush)
        for h in range(25):
            x = h/24*W
            p.drawLine(x,y+barH+2,x,y+barH+6)
            if h%4==0:
                p.setPen(QColor("#666")); p.setFont(QFont("Arial",7))
                p.drawText(x-10,y+barH+14, f"{h:02d}")
                p.setPen(QColor("#444"))


#─────────────────────────  WEEKLY CHART  ────────────────────────────────────
class WeeklyChartWidget(QWidget):
    """Bar chart showing work hours for each day of the week."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.data = {}  # Dictionary mapping day -> hours
        self.days_of_week = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
        self.setMinimumSize(500, 250)
    
    def setData(self, data: dict[str, float]):
        self.data = data
        self.update()
    
    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        W, H = self.width(), self.height()
        
        # Title
        p.setPen(QColor("#CDD6F4"))
        p.setFont(QFont("Arial", 14, QFont.Bold))
        p.drawText(QRectF(0, 0, W, 30), Qt.AlignCenter, "Work Hours by Day")
        
        # Chart area
        chartTop = 50
        chartBottom = H - 40
        chartHeight = chartBottom - chartTop
        
        # Calculate max hours and average
        values = list(self.data.values())
        maxHours = max(values) if values else 0
        maxHours = max(8.0, maxHours)  # Ensure at least 8 hours scale
        avgHours = sum(values) / len(values) if values else 0
        
        # Y-axis labels and grid lines
        p.setPen(QColor("#444"))
        p.setFont(QFont("Arial", 9))
        for i in range(5):  # 0, 2, 4, 6, 8
            value = i * 2
            y = chartBottom - (value / maxHours) * chartHeight
            p.drawLine(50, y, W - 50, y)  # Grid line
            p.drawText(30, y + 5, str(value))  # Y-axis label
        
        # Average line
        if values:
            avgY = chartBottom - (avgHours / maxHours) * chartHeight
            p.setPen(QPen(QColor("#FFF"), 1, Qt.DashLine))
            p.drawLine(50, avgY, W - 50, avgY)
            p.setPen(QColor("#FFF"))
            p.drawText(W - 80, avgY - 10, "Average")
            p.drawText(W - 80, avgY + 10, f"{avgHours:.1f}")
        
        # X-axis and bars
        barWidth = (W - 100) / 7
        for i, day in enumerate(self.days_of_week):
            x = 50 + i * barWidth
            
            # Draw day label
            p.setPen(QColor("#888"))
            p.drawText(QRectF(x, chartBottom + 10, barWidth, 20), Qt.AlignCenter, day)
            
            # Draw bar if we have data
            hours = self.data.get(day, 0)
            if hours > 0:
                barHeight = (hours / maxHours) * chartHeight
                barRect = QRectF(x + barWidth * 0.2, chartBottom - barHeight, barWidth * 0.6, barHeight)
                p.setPen(Qt.NoPen)
                p.setBrush(QBrush(CLR_FOCUS))
                p.drawRect(barRect)
                
                # Draw hour value above bar
                p.setPen(QColor("#FFF"))
                p.drawText(QRectF(x, chartBottom - barHeight - 20, barWidth, 20), 
                          Qt.AlignCenter, f"{hours:.1f}")


#─────────────────────────  MONTHLY CHART  ────────────────────────────────────
class MonthlyChartWidget(QWidget):
    """Chart showing work hours for each week in a month."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.data = []  # List of (week_num, date_range, hours) tuples
        self.setMinimumSize(500, 250)
    
    def setData(self, data: list[tuple[int, str, float]]):
        self.data = data
        self.update()
    
    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        W, H = self.width(), self.height()
        
        # Title
        p.setPen(QColor("#CDD6F4"))
        p.setFont(QFont("Arial", 14, QFont.Bold))
        p.drawText(QRectF(0, 0, W, 30), Qt.AlignCenter, "Weekly Work Hours")
        
        if not self.data:
            p.setPen(QColor("#666"))
            p.setFont(QFont("Arial", 10))
            p.drawText(self.rect(), Qt.AlignCenter, "No Data")
            return
        
        # Calculate max hours
        maxHours = max(week[2] for week in self.data) if self.data else 0
        maxHours = max(40.0, maxHours)  # Ensure at least 40 hours scale for weekly view
        
        # Draw weekly bars
        barHeight = 30
        spacing = 15
        startY = 60
        
        for i, (week_num, date_range, hours) in enumerate(self.data):
            y = startY + i * (barHeight + spacing)
            
            # Week label
            p.setPen(QColor("#CDD6F4"))
            p.setFont(QFont("Arial", 9))
            p.drawText(10, y + barHeight/2 + 5, f"Week {week_num}")
            
            # Date range
            p.setPen(QColor("#888"))
            p.setFont(QFont("Arial", 8))
            p.drawText(10, y + barHeight/2 + 20, date_range)
            
            # Progress bar background
            barRect = QRectF(100, y, W - 150, barHeight)
            p.setPen(Qt.NoPen)
            p.setBrush(QBrush(QColor("#272738")))
            p.drawRoundedRect(barRect, 4, 4)
            
            # Progress bar fill
            fillWidth = (hours / maxHours) * (W - 150)
            fillRect = QRectF(100, y, fillWidth, barHeight)
            p.setBrush(QBrush(CLR_FOCUS))
            p.drawRoundedRect(fillRect, 4, 4)
            
            # Hours text
            p.setPen(QColor("#CDD6F4"))
            p.setFont(QFont("Arial", 9))
            p.drawText(W - 40, y + barHeight/2 + 5, f"{hours:.1f}h")


#─────────────────────────  CATEGORY BAR WIDGET  ──────────────────────────────
class CategoryBarWidget(QWidget):
    def __init__(self,label,secs,max_secs,color,parent=None):
        super().__init__(parent); self.label=label; self.secs=secs; self.max=max_secs; self.color=color
        self.setFixedHeight(24); self.setToolTip(self.fmt(secs))
    def fmt(self,s): h,m=divmod(int(s//60),60); return f"{h}h {m}m" if h else f"{m}m"
    def paintEvent(self,_):
        p=QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        w,h=self.width(),self.height(); bw=w*0.7*self.secs/self.max if self.max else 0
        p.setPen(Qt.NoPen); p.setBrush(self.color)
        p.drawRoundedRect(QRectF(0,4,bw,h-8),4,4)
        p.setPen(QColor("#CDD6F4")); p.setFont(QFont("Arial",9))
        txt=f"{self.label[:14]}{'…' if len(self.label)>14 else ''} ({self.fmt(self.secs)})"
        p.drawText(QRectF(w*0.72,0,w*0.28,h),Qt.AlignLeft|Qt.AlignVCenter,txt)


#─────────────────────────  DASHBOARD DIALOG  ────────────────────────────────
class DashboardWindow(QDialog):
    def __init__(self,parent=None,log_dir:str|Path="."):
        super().__init__(parent)
        self.log_dir=Path(log_dir)
        self.mode="day"; self.date=QDate.currentDate()

        self.setMinimumSize(880,680); self.setWindowTitle("FocusTimer Dashboard")
        self.setStyleSheet("""
            QDialog{background:#1E1E2E;color:#CDD6F4;}
            QLabel{color:#CDD6F4;}
            QFrame#card{background:#28283D;border-radius:8px;}
            QPushButton{background:#313244;color:#CDD6F4;border:none;
                        padding:4px 10px;border-radius:4px;min-width:60px;}
            QPushButton:hover{background:#45475A;}
            QPushButton:checked{background:#89B4FA;color:#1E1E2E;}
            QDateEdit{background:#313244;color:#CDD6F4;border:1px solid #555;
                      padding:3px;border-radius:3px;}
        """)
        self.buildUI(); self.load()

    # UI ------------------------------------------------------------------
    def buildUI(self):
        root=QVBoxLayout(self); root.setContentsMargins(20,20,20,20); root.setSpacing(18)
        # header
        hdr=QHBoxLayout(); root.addLayout(hdr)
        hdr.addWidget(QLabel("Productivity Dashboard",font=QFont("Arial",18,QFont.Bold)))
        hdr.addStretch(1)
        self.dayB  = self.mkModeBtn("Day")
        self.weekB = self.mkModeBtn("Week", checked=False)
        self.monthB= self.mkModeBtn("Month",checked=False)
        for b in (self.dayB,self.weekB,self.monthB): hdr.addWidget(b)
        hdr.addSpacing(15)
        self.prevB=QPushButton("←"); self.prevB.setFixedWidth(28)
        self.nextB=QPushButton("→"); self.nextB.setFixedWidth(28)
        self.dateE=QDateEdit(self.date); self.dateE.setCalendarPopup(True)
        hdr.addWidget(self.prevB); hdr.addWidget(self.dateE); hdr.addWidget(self.nextB)
        self.prevB.clicked.connect(lambda: self.shift(-1))
        self.nextB.clicked.connect(lambda: self.shift(+1))
        self.dateE.dateChanged.connect(self.onDate)
        # top summary
        top=QHBoxLayout(); top.setSpacing(15); root.addLayout(top)
        card=QFrame(); card.setObjectName("card")
        cLay=QVBoxLayout(card); cLay.setContentsMargins(15,15,15,15)
        cLay.addWidget(QLabel("Summary",font=QFont("Arial",14,QFont.Bold)))
        self.workL=QLabel(); self.focusL=QLabel()
        cLay.addWidget(self.workL); cLay.addWidget(self.focusL); cLay.addStretch()
        top.addWidget(card,1)
        # donut
        dCard=QFrame(); dCard.setObjectName("card")
        dLay=QVBoxLayout(dCard)
        dLay.addWidget(QLabel("Time Breakdown",alignment=Qt.AlignCenter,
                               font=QFont("Arial",14,QFont.Bold)))
        self.donut=DonutChartWidget(); dLay.addWidget(self.donut)
        top.addWidget(dCard,1)
        
        # Middle content area (Day/Week/Month specific views)
        self.contentArea = QVBoxLayout()
        root.addLayout(self.contentArea)
        
        # Timeline (Day view)
        self.timelineCard = QFrame()
        self.timelineCard.setObjectName("card")
        tLay = QVBoxLayout(self.timelineCard)
        tLay.addWidget(QLabel("Calendar (sessions)", font=QFont("Arial", 14, QFont.Bold)))
        self.timeline = TimelineWidget()
        tLay.addWidget(self.timeline)
        
        # Weekly chart (Week view)
        self.weeklyCard = QFrame()
        self.weeklyCard.setObjectName("card")
        wLay = QVBoxLayout(self.weeklyCard)
        wLay.addWidget(QLabel("Daily Breakdown", font=QFont("Arial", 14, QFont.Bold)))
        self.weeklyChart = WeeklyChartWidget()
        wLay.addWidget(self.weeklyChart)
        
        # Monthly chart (Month view)
        self.monthlyCard = QFrame()
        self.monthlyCard.setObjectName("card")
        mLay = QVBoxLayout(self.monthlyCard)
        mLay.addWidget(QLabel("Weekly Breakdown", font=QFont("Arial", 14, QFont.Bold)))
        self.monthlyChart = MonthlyChartWidget()
        mLay.addWidget(self.monthlyChart)
        
        # Initially add the timeline card (day view)
        self.contentArea.addWidget(self.timelineCard)
        self.weeklyCard.hide()
        self.monthlyCard.hide()
        
        # categories
        cat=QFrame(); cat.setObjectName("card")
        bLay=QVBoxLayout(cat); bLay.setContentsMargins(15,15,15,15)
        bLay.addWidget(QLabel("Top Categories",font=QFont("Arial",14,QFont.Bold)))
        self.catBox=QVBoxLayout()
        scr=QScrollArea(); scr.setWidgetResizable(True)
        cont=QWidget(); cont.setLayout(self.catBox); scr.setWidget(cont)
        bLay.addWidget(scr); bLay.addStretch()
        root.addWidget(cat)

    def mkModeBtn(self,txt,checked=True):
        b=QPushButton(txt); b.setCheckable(True); b.setChecked(checked)
        b.clicked.connect(lambda _=False,m=txt.lower(): self.setMode(m)); return b

    # mode / date helpers --------------------------------------------------
    def setMode(self,m):
        if m==self.mode: return
        self.mode=m
        for b,name in ((self.dayB,"day"),(self.weekB,"week"),(self.monthB,"month")):
            b.setChecked(name==m)
        fmt={"day":"yyyy-MM-dd","week":"yyyy 'W'ww","month":"MMM yyyy"}[m]
        self.dateE.setDisplayFormat(fmt)
        
        # Update content area based on mode
        self.timelineCard.hide()
        self.weeklyCard.hide()
        self.monthlyCard.hide()
        
        if m == "day":
            self.contentArea.addWidget(self.timelineCard)
            self.timelineCard.show()
        elif m == "week":
            self.contentArea.addWidget(self.weeklyCard)
            self.weeklyCard.show()
        elif m == "month":
            self.contentArea.addWidget(self.monthlyCard)
            self.monthlyCard.show()
            
        self.load()

    def shift(self,d):
        date=self.dateE.date()
        if self.mode=="day":   date=date.addDays(d)
        elif self.mode=="week":date=date.addDays(7*d)
        else:                  date=date.addMonths(d)
        if date<=QDate.currentDate(): self.dateE.setDate(date)

    def onDate(self,d): self.date=d; self.load()

    # ----------------------------------------------------------------------
    def load(self):
        # clear UI
        self.workL.setText("…"); self.focusL.clear(); self.donut.setData({})
        while self.catBox.count():
            w=self.catBox.takeAt(0).widget()
            if w: w.deleteLater()
        self.timeline.setData([])

        dates = self.collectDates()
        totHrs=focusS=breakS=0; apps=defaultdict(int)
        # needed for timeline
        intervals=[]
        
        # Weekly data
        weeklyData = {day: 0.0 for day in ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']}
        
        # Monthly data (week number -> hours)
        monthlyData = []
        
        for d in dates:
            f=self.log_dir / d.toString("yyyy-MM-dd") / "activity_log.csv"
            if not f.exists(): continue
            try:
                with f.open() as csvfile:
                    rows=list(csv.reader(csvfile))
                if len(rows)<=1: continue
                try: totHrs+=float(rows[-1][3])
                except: pass
                prev_ts=None
                for r in rows[1:]:
                    if len(r)<5: continue
                    ts=datetime.strptime(r[0],"%Y-%m-%d %H:%M:%S")
                    active= int(r[1]) if r[1].isdigit() else 0
                    app   = r[4] or "Unknown"
                    if prev_ts:
                        dur=(ts-prev_ts).total_seconds()
                        if dur>0:
                            if active:
                                focusS+=dur; apps[app]+=dur
                            else:
                                breakS+=dur; apps["Break"]+=dur
                            if self.mode=="day":
                                intervals.append((prev_ts,ts, bool(active)))
                            
                            # Update weekly data
                            if self.mode=="week":
                                day_name = prev_ts.strftime("%a")
                                weeklyData[day_name] += dur / 3600  # Convert seconds to hours
                    prev_ts=ts
                    
                # For month view, calculate total hours for the week
                if self.mode=="month":
                    week_num = d.weekNumber()[0]
                    week_start = d.addDays(-d.dayOfWeek() + 1)  # Monday of the week
                    week_end = week_start.addDays(6)  # Sunday of the week
                    date_range = f"{week_start.day()}/{week_start.month()} - {week_end.day()}/{week_end.month()}"
                    
                    # Check if we already have this week in our data
                    week_exists = False
                    for i, (w_num, _, hours) in enumerate(monthlyData):
                        if w_num == week_num:
                            monthlyData[i] = (w_num, date_range, monthlyData[i][2] + totHrs)
                            week_exists = True
                            break
                    
                    if not week_exists and totHrs > 0:
                        monthlyData.append((week_num, date_range, totHrs))
            except Exception as e:
                print("error", f,e); continue

        # Summary labels
        h,m=divmod(int(totHrs*60),60); self.workL.setText(f"Work Hours: {h} h {m} m")
        totS=focusS+breakS
        self.focusL.setText(f"Focus: {(focusS/totS*100):.0f}%" if totS else "Focus: 0%")
        self.donut.setData({"Focus":focusS,"Break":breakS})

        # Update view-specific data
        if self.mode=="day":
            self.timeline.setData(intervals)
        elif self.mode=="week":
            self.weeklyChart.setData(weeklyData)
        elif self.mode=="month":
            # Sort monthly data by week number
            monthlyData.sort(key=lambda x: x[0])
            self.monthlyChart.setData(monthlyData)

        if apps:
            pal=["#F2CDCD","#DDB6F2","#F5C2E7","#E8A2AF","#F28FAD","#ABE9B3",
                 "#FAE3B0","#F8BD96","#EE99A0","#89DCEB"]
            top10=sorted(apps.items(),key=lambda x:x[1],reverse=True)[:10]
            max_s=top10[0][1]
            for i,(name,secs) in enumerate(top10):
                self.catBox.addWidget(CategoryBarWidget(name,secs,max_s,QColor(pal[i%len(pal)])))
        else:
            self.catBox.addWidget(QLabel("No data available."))

    def collectDates(self):
        if self.mode=="day": return [self.date]
        if self.mode=="week":
            mon=self.date.addDays(1-self.date.dayOfWeek())
            return [mon.addDays(i) for i in range(7)
                    if mon.addDays(i)<=QDate.currentDate()]
        # month
        days=self.date.daysInMonth()
        return [QDate(self.date.year(),self.date.month(),d)
                for d in range(1,days+1)
                if QDate(self.date.year(),self.date.month(),d)<=QDate.currentDate()]


#─────────────────────────  manual run  ──────────────────────────────────────
if __name__=="__main__":
    logbase=Path(sys.argv[1] if len(sys.argv)>1 else "./focus_logs")
    app=QApplication(sys.argv); dlg=DashboardWindow(log_dir=logbase); dlg.show()
    sys.exit(app.exec_())