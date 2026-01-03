import os
import sys
import json
import sqlite3
import datetime
import csv
import traceback
import requests
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk, filedialog
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib.dates as mdates

# --- CONFIGURAZIONE ---
SETTINGS_FILE = "meteo_settings.json"
DB_FILE = "cantiere_history.db"

DEFAULT_SETTINGS = {
    "api_key": "10cd392e7b7ac7e0d18694484f80977d",
    "city": "Priolo Gargallo,IT",
    "language": "it",
    "thresholds": {
        "rain_mm_3h": 2.0, "wind_kmh": 40.0, "temp_min_c": 5.0,
        "temp_max_c": 30.0, "visibility_min_m": 2000, "humidity_max_pct": 85
    }
}

class SettingsManager:
    @staticmethod
    def load():
        if not os.path.exists(SETTINGS_FILE): return DEFAULT_SETTINGS.copy()
        try:
            with open(SETTINGS_FILE, 'r') as f: return json.load(f)
        except: return DEFAULT_SETTINGS.copy()
    @staticmethod
    def save(data):
        try:
            with open(SETTINGS_FILE, 'w') as f: json.dump(data, f, indent=4)
            return True
        except: return False

# --- DATABASE (REGISTRATORE DI DATI) ---
class DatabaseManager:
    def __init__(self):
        self.conn = sqlite3.connect(DB_FILE)
        # Tabella Giornaliera (Sintesi)
        self.conn.execute('''CREATE TABLE IF NOT EXISTS daily_history (
            date_str TEXT, city TEXT, status TEXT, operability INTEGER,
            min_temp REAL, max_temp REAL, max_wind REAL, total_rain REAL,
            min_visibility INTEGER, max_humidity INTEGER, alert_notes TEXT,
            PRIMARY KEY (date_str, city))''')
        # Tabella Oraria (Registratore Real-Time)
        self.conn.execute('''CREATE TABLE IF NOT EXISTS hourly_history (
            timestamp DATETIME, city TEXT, temp REAL, wind REAL, rain REAL, hum INTEGER, desc TEXT,
            PRIMARY KEY (timestamp, city))''')
        self.conn.commit()

    def save_snapshot(self, city, data):
        """Salva il meteo attuale per lo storico delle ore precedenti"""
        now = datetime.datetime.now().replace(minute=0, second=0, microsecond=0)
        self.conn.execute('''INSERT OR REPLACE INTO hourly_history VALUES (?,?,?,?,?,?,?)''',
            (now, city, data['temp'], data['wind'], 0.0, data['hum'], data['desc']))
        self.conn.commit()

    def get_today_snapshots(self, city):
        """Recupera i dati reali registrati oggi"""
        today = datetime.date.today().isoformat()
        cursor = self.conn.execute('''SELECT timestamp, temp, wind, rain, hum, desc FROM hourly_history 
                                   WHERE city = ? AND timestamp >= ? ORDER BY timestamp ASC''', (city, today))
        return cursor.fetchall()

    def upsert_day(self, city, status, operability, min_t, max_t, max_w, tot_r, min_vis, max_hum, alerts):
        today_str = datetime.date.today().strftime("%d/%m/%Y")
        notes = ", ".join(alerts) if alerts else ""
        self.conn.execute('''INSERT OR REPLACE INTO daily_history VALUES (?,?,?,?,?,?,?,?,?,?,?)''',
            (today_str, city, status, operability, min_t, max_t, max_w, tot_r, min_vis, max_hum, notes))
        self.conn.commit()

    def get_history(self, city=None):
        query = 'SELECT * FROM daily_history'
        if city: query += f" WHERE city = '{city}'"
        query += ' ORDER BY substr(date_str, 7, 4) DESC, substr(date_str, 4, 2) DESC, substr(date_str, 1, 2) DESC LIMIT 100'
        return self.conn.execute(query).fetchall()

    def export_csv(self, filename):
        cursor = self.conn.execute("SELECT * FROM daily_history")
        rows = cursor.fetchall()
        headers = ["Data", "Città", "Stato", "Op%", "MinT", "MaxT", "Vento", "Pioggia", "Visib", "Umid", "Note"]
        with open(filename, 'w', newline='', encoding='utf-8') as f:
            w = csv.writer(f, delimiter=';'); w.writerow(headers); w.writerows(rows)

class MeteoAPI:
    def __init__(self, settings):
        self.settings = settings; self.lat = None; self.lon = None; self.city_name = None
    def get_data(self):
        g_url = f"http://api.openweathermap.org/geo/1.0/direct?q={self.settings['city']}&limit=1&appid={self.settings['api_key']}"
        g_data = requests.get(g_url, timeout=10).json()
        if not g_data: raise Exception("Città non trovata")
        self.lat, self.lon, self.city_name = g_data[0]['lat'], g_data[0]['lon'], g_data[0]['name']
        c_url = f"http://api.openweathermap.org/data/2.5/weather?lat={self.lat}&lon={self.lon}&appid={self.settings['api_key']}&units=metric&lang={self.settings['language']}"
        f_url = f"http://api.openweathermap.org/data/2.5/forecast?lat={self.lat}&lon={self.lon}&appid={self.settings['api_key']}&units=metric&lang={self.settings['language']}"
        return requests.get(c_url, timeout=10).json(), requests.get(f_url, timeout=10).json()

class DataProcessor:
    @staticmethod
    def get_icon(desc):
        d = desc.lower()
        if "clear" in d or "sereno" in d: return "☀️"
        if "cloud" in d or "nuvol" in d: return "☁️"
        if "rain" in d or "pioggia" in d: return "🌧️"
        if "storm" in d or "temporale" in d: return "⚡"
        return "🌥️"

    @staticmethod
    def analyze(current, forecast, thresholds, past_snapshots):
        # 1. Dati Attuali
        now_data = {
            'temp': current['main']['temp'], 'wind': current['wind']['speed']*3.6,
            'hum': current['main']['humidity'], 'vis': current.get('visibility', 10000),
            'desc': current['weather'][0]['description'].capitalize(),
            'icon': DataProcessor.get_icon(current['weather'][0]['description']),
            'sunrise': datetime.datetime.fromtimestamp(current['sys']['sunrise']).strftime('%H:%M'),
            'sunset': datetime.datetime.fromtimestamp(current['sys']['sunset']).strftime('%H:%M')
        }

        # 2. Timeline Unificata (Passato + Futuro)
        timeline = []
        # Aggiungi Passato (dal DB)
        for p in past_snapshots:
            ts = datetime.datetime.fromisoformat(p[0])
            if ts.hour < datetime.datetime.now().hour:
                timeline.append({'dt': ts.timestamp(), 'temp': p[1], 'wind': p[2], 'rain': 0.0, 'hum': p[4], 'desc': p[5], 'type': 'RILEVATO'})
        
        # Aggiungi Futuro (dalle previsioni)
        for f in forecast['list'][:10]:
            timeline.append({'dt': f['dt'], 'temp': f['main']['temp'], 'wind': f['wind']['speed']*3.6, 'rain': f.get('rain',{}).get('3h',0), 'desc': f['weather'][0]['description'], 'type': 'PREVISTO'})

        summary = {
            'min_temp': 99, 'max_temp': -99, 'max_wind': 0, 'total_rain': 0,
            'status': 'OK', 'score': 100, 'alerts': [], 'safety': [], 'current': now_data, 'timeline': timeline
        }
        
        # Logica Score e Alerts
        risk = 0
        for i in timeline:
            if i['type'] == 'PREVISTO':
                if i['temp'] > thresholds['temp_max_c']: risk += 5
                if i['wind'] > thresholds['wind_kmh']: risk += 15
                if i.get('rain',0) > thresholds['rain_mm_3h']: summary['status']="CRITICO"; risk += 40
        
        summary['score'] = max(0, 100 - risk)
        return summary

class PDFReporter:
    @staticmethod
    def generate(city, summary, base_dir):
        path = os.path.join(base_dir, "REPORT_OUTPUT"); os.makedirs(path, exist_ok=True)
        fname = os.path.join(path, f"Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.pdf")
        doc = SimpleDocTemplate(fname, pagesize=A4); elems = []; styles = getSampleStyleSheet()
        elems.append(Paragraph(f"DIARIO METEO CANTIERE - {city}", styles['Title']))
        elems.append(Paragraph(f"Timeline completa (Rilevazioni + Previsioni)", styles['Normal']))
        rows = [['Data/Ora', 'Tipo', 'Meteo', 'T°C', 'Vento']]
        for i in summary['timeline']:
            dt = datetime.datetime.fromtimestamp(i['dt']).strftime('%d/%m/%Y %H:%M')
            rows.append([dt, i['type'], i['desc'], f"{i['temp']:.1f}", f"{i['wind']:.0f}"])
        t = Table(rows); t.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.5,colors.grey),('BACKGROUND',(0,0),(-1,0),colors.navy),('TEXTCOLOR',(0,0),(-1,0),colors.white)]))
        elems.append(t); doc.build(elems); return fname

class MeteoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ISAB Manager v7.3 - Black Box")
        # Apertura massimizzata (Schermo Intero con barra applicazioni)
        self.root.state('zoomed') 
        self.root.configure(bg="#eceff1")
        self.db = DatabaseManager(); self.settings = SettingsManager.load(); self.api = MeteoAPI(self.settings)
        self._setup_ui(); self.root.after(100, self.refresh)

    def _setup_ui(self):
        self.tabs = ttk.Notebook(self.root); self.tabs.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.tab_dash = tk.Frame(self.tabs, bg="#eceff1"); self.tab_hist = tk.Frame(self.tabs, bg="white"); self.tab_sett = tk.Frame(self.tabs, bg="white")
        self.tabs.add(self.tab_dash, text=" ⚡ DASHBOARD "); self.tabs.add(self.tab_hist, text=" 📊 STORICO GIORNALIERO "); self.tabs.add(self.tab_sett, text=" ⚙️ SETTINGS ")
        
        # Dashboard UI
        self.hdr = tk.Frame(self.tab_dash, bg="#263238", height=90); self.hdr.pack(fill=tk.X)
        self.lbl_now = tk.Label(self.hdr, text="--°C", font=("Segoe UI", 24, "bold"), bg="#263238", fg="white"); self.lbl_now.pack(side=tk.LEFT, padx=20)
        self.lbl_status = tk.Label(self.hdr, text="STATO: ---", font=("Segoe UI", 16, "bold"), bg="#263238", fg="white"); self.lbl_status.pack(side=tk.RIGHT, padx=20)

        paned = ttk.PanedWindow(self.tab_dash, orient=tk.HORIZONTAL); paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        left = tk.Frame(paned, bg="#eceff1"); paned.add(left, weight=1)
        self.log = scrolledtext.ScrolledText(left, font=("Consolas", 10), state='disabled'); self.log.pack(fill=tk.BOTH, expand=True)
        self.log.tag_config("PAST", foreground="gray"); self.log.tag_config("NOW", background="#e3f2fd", font=("bold")); self.log.tag_config("FUTURE", foreground="black")

        self.fig = Figure(figsize=(5, 5), dpi=100, facecolor='#eceff1'); self.ax = self.fig.add_subplot(111); self.cv = FigureCanvasTkAgg(self.fig, master=paned); paned.add(self.cv.get_tk_widget(), weight=2)
        
        ft = tk.Frame(self.tab_dash, bg="#eceff1"); ft.pack(fill=tk.X, side=tk.BOTTOM, pady=5)
        tk.Button(ft, text="AGGIORNA", command=self.refresh, bg="#2196f3", fg="white", width=20).pack(side=tk.LEFT, padx=10)
        self.btn_pdf = tk.Button(ft, text="REPORT PDF", command=self.make_pdf, state='disabled', bg="#4caf50", fg="white", width=20); self.btn_pdf.pack(side=tk.RIGHT, padx=10)

    def refresh(self):
        try:
            curr, fore = self.api.get_data()
            self.db.save_snapshot(self.api.city_name, {'temp': curr['main']['temp'], 'wind': curr['wind']['speed']*3.6, 'hum': curr['main']['humidity'], 'desc': curr['weather'][0]['description']})
            past = self.db.get_today_snapshots(self.api.city_name)
            self.res = DataProcessor.analyze(curr, fore, self.settings['thresholds'], past)
            self.update_ui(); self.btn_pdf.config(state='normal')
        except Exception as e: messagebox.showerror("Err", str(e))

    def update_ui(self):
        r = self.res; cur = r['current']
        self.lbl_now.config(text=f"{cur['icon']} {cur['temp']:.1f}°C")
        col = "#4caf50" if r['status'] == "OK" else "#ffa000" if r['status'] == "ATTENZIONE" else "#d32f2f"
        self.hdr.config(bg=col); self.lbl_status.config(text=f"STATO: {r['status']} ({r['score']} %)", bg=col)

        self.log.config(state='normal'); self.log.delete('1.0', tk.END)
        self.log.insert(tk.END, "--- TIMELINE DIARIO (RILEVATO + PREVISTO) ---\n\n", "DATE")
        
        for i in r['timeline']:
            dt = datetime.datetime.fromtimestamp(i['dt'])
            tag = "PAST" if i['type']=='RILEVATO' else "FUTURE"
            if dt.hour == datetime.datetime.now().hour and dt.date() == datetime.date.today(): tag = "NOW"
            prefix = "●" if i['type']=='RILEVATO' else "○"
            self.log.insert(tk.END, f"{prefix} {dt.strftime('%d/%m/%Y %H:%M')} [{i['type']}] {i['temp']:.1f}°C - {i['desc']}\n", tag)
        self.log.config(state='disabled')

        # Chart configuration
        self.ax.clear()
        times = [datetime.datetime.fromtimestamp(x['dt']) for x in r['timeline']]
        temps = [x['temp'] for x in r['timeline']]
        
        # Sfondo tratteggiato dell'andamento generale
        self.ax.plot(times, temps, 'b--', alpha=0.2) 
        
        # Divisione Rilevato/Previsto
        now_idx = len([x for x in r['timeline'] if x['type']=='RILEVATO'])
        self.ax.plot(times[:now_idx+1], temps[:now_idx+1], 'r-o', linewidth=2, label='Storico Rilevato')
        self.ax.plot(times[now_idx:], temps[now_idx:], 'g-o', linewidth=2, label='Previsione Futura')
        
        # Linea verticale del tempo attuale
        self.ax.axvline(datetime.datetime.now(), color='black', linestyle=':', alpha=0.5) 

        # Titoli e Coordinate (X e Y)
        self.ax.set_title("DIARIO TERMICO CANTIERE (24h)", fontsize=12, fontweight='bold', pad=15)
        self.ax.set_xlabel("Data e Ora del Rilevamento", fontsize=10, labelpad=10)
        self.ax.set_ylabel("Temperatura misurata in °C", fontsize=10, labelpad=10)
        
        # Formattatore asse X per GG/MM/AAAA HH:MM
        self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m/%Y\n%H:%M'))
        
        self.ax.legend(loc='upper right', fontsize=9)
        self.ax.grid(True, linestyle=':', alpha=0.6)
        
        # Formattazione asse X (Date/Ore)
        self.fig.autofmt_xdate()
        self.cv.draw()

    def make_pdf(self):
        base = os.path.dirname(os.path.abspath(__file__)); p = PDFReporter.generate(self.api.city_name, self.res, base); os.startfile(p)

sys.excepthook = lambda t, v, tb: open("crash.log","a").write(traceback.format_exc())
if __name__ == "__main__": root = tk.Tk(); MeteoApp(root); root.mainloop()