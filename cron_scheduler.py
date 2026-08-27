"""
Background Cron Scheduler Engine for Instagram Analytics
Can run in background threads alongside Streamlit or as a standalone daemon.
"""

import sys, os, json, time, threading
from datetime import datetime, timezone
import schedule

sys.stdout.reconfigure(encoding="utf-8")

CRON_CONFIG_FILE = "cron_config.json"
CRON_STATUS_FILE = "cron_status.json"

DEFAULT_CONFIG = {
    "enabled": True,
    "schedule_type": "daily",     # "daily", "hourly", "weekly", "interval_hours"
    "scheduled_time": "09:00",    # Time for daily / weekly
    "interval_hours": 24,         # For interval mode
    "weekday": "monday",          # For weekly mode
    "brands": [
        {"name": "GRT Oriana", "handle": "grtoriana", "category": "Jewellery"},
        {"name": "Gully Labs", "handle": "gullylabs", "category": "Footwear"},
        {"name": "Skechers India", "handle": "skechersindia", "category": "Footwear"},
        {"name": "Japam", "handle": "japam.in", "category": "Spiritual"},
        {"name": "Divine Hindu", "handle": "divinehindu.in", "category": "Spiritual"},
        {"name": "Comet", "handle": "thecometuniverse", "category": "Footwear"}
    ],
    "google_sheets_webhook_url": "", # Optional: push to Google Sheets
    "last_run": None,
    "next_run": None,
    "run_count": 0
}

def load_cron_config():
    if os.path.exists(CRON_CONFIG_FILE):
        try:
            with open(CRON_CONFIG_FILE, "r", encoding="utf-8") as f:
                return {**DEFAULT_CONFIG, **json.load(f)}
        except Exception:
            pass
    return DEFAULT_CONFIG

def save_cron_config(cfg):
    with open(CRON_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)

def update_cron_status(status_msg, is_running=False, error=None):
    status = {
        "is_running": is_running,
        "last_update": datetime.now(timezone.utc).isoformat(),
        "status_message": status_msg,
        "error": error
    }
    with open(CRON_STATUS_FILE, "w", encoding="utf-8") as f:
        json.dump(status, f, indent=2)

def execute_scheduled_brand_audit():
    """Main task executed by the cron scheduler"""
    cfg = load_cron_config()
    print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ⏰ CRON TRIGGER: Starting scheduled brand audit...", flush=True)
    update_cron_status("Running scheduled brand audit...", is_running=True)
    
    try:
        from instagram_paid_collabs_api import InstagramPaidPartnershipEngine
        engine = InstagramPaidPartnershipEngine(max_workers=10)
        
        all_results = []
        for b in cfg["brands"]:
            h = b["handle"].replace("@", "").strip()
            print(f"  -> Auditing brand: @{h}...", flush=True)
            res = engine.analyze_brand(h, max_pages=6)
            all_results.append(res)
            time.sleep(1.0)
            
        # Update config metrics
        cfg["last_run"] = datetime.now(timezone.utc).isoformat()
        cfg["run_count"] = cfg.get("run_count", 0) + 1
        save_cron_config(cfg)
        
        # Save snapshot
        with open("latest_cron_audit_snapshot.json", "w", encoding="utf-8") as f:
            json.dump(all_results, f, indent=2)
            
        print(f"✅ CRON SUCCESS: Audited {len(all_results)} brands successfully!\n", flush=True)
        update_cron_status(f"Completed audit for {len(all_results)} brands at {datetime.now().strftime('%H:%M:%S')}", is_running=False)
        
    except Exception as e:
        print(f"❌ CRON ERROR: {e}", flush=True)
        update_cron_status(f"Error during audit: {str(e)}", is_running=False, error=str(e))

class BackgroundCronDaemon:
    """Thread-safe background scheduler daemon"""
    def __init__(self):
        self.thread = None
        self.running = False
        
    def start(self):
        if self.running:
            return
        self.running = True
        self.thread = threading.Thread(target=self._run_loop, daemon=True)
        self.thread.start()
        print("✓ Background Cron Scheduler daemon started!")
        
    def stop(self):
        self.running = False
        schedule.clear()
        
    def _run_loop(self):
        while self.running:
            schedule.run_pending()
            time.sleep(1)

_daemon = None

def get_or_start_cron_daemon():
    global _daemon
    if _daemon is None:
        _daemon = BackgroundCronDaemon()
        cfg = load_cron_config()
        schedule.clear()
        
        if cfg.get("enabled", True):
            stype = cfg.get("schedule_type", "daily")
            stime = cfg.get("scheduled_time", "09:00")
            
            if stype == "daily":
                schedule.every().day.at(stime).do(execute_scheduled_brand_audit)
            elif stype == "hourly":
                schedule.every().hour.do(execute_scheduled_brand_audit)
            elif stype == "weekly":
                wday = cfg.get("weekday", "monday").lower()
                getattr(schedule.every(), wday).at(stime).do(execute_scheduled_brand_audit)
            elif stype == "interval_hours":
                schedule.every(cfg.get("interval_hours", 24)).hours.do(execute_scheduled_brand_audit)
                
        _daemon.start()
    return _daemon

if __name__ == "__main__":
    print("Running standalone cron scheduler test...")
    execute_scheduled_brand_audit()
