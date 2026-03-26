# UGC Creator Sourcing Pipeline

A complete system for sourcing and vetting creators for UGC campaigns. Scans Twitter/X profiles and scrapes the Kaito leaderboard to build enriched creator lists.

---

## What's in the Box

| Script | Purpose | Use For |
|--------|---------|---------|
| `twitter_scanner_v2.py` | Full Twitter/X profile scanner | Your existing xlsx spreadsheets (both sheets) |
| `verified_only_scanner.py` | Lightweight verified-only check | Large batches where you only need verification status |
| `kaito_pipeline.py` | Kaito leaderboard scraper + Twitter enrichment | Sourcing top crypto/Web3 creators |
| `scan_both_sheets.py` | Batch runner for both xlsx sheets | Convenience script |
| `twitter_scanner.py` | Original v1 scanner (legacy) | Basic scan without country/following |

---

## Data Collected

### Twitter Scanner v2 (full enrichment)
| Field | Description |
|-------|-------------|
| Username | Twitter/X handle |
| Blue Check Verified | Yes / No |
| Followers | Human-readable (e.g. 12.5K) |
| Followers (Raw) | Exact number |
| Following | Human-readable |
| Following (Raw) | Exact number |
| Follower/Following Ratio | Engagement quality signal |
| Country / Region | Self-reported location from profile |
| Campaign Ready | Auto-flag: Yes if 1,000+ followers |
| Status | OK, Not Found, Private, Error |

### Kaito Pipeline (additional fields)
| Field | Description |
|-------|-------------|
| Rank | Position on Kaito leaderboard |
| Yaps Score | Kaito influence score (optional enrichment) |
| + All Twitter fields above | |

---

## Setup (One-Time)

```bash
cd twitter_scanner
pip install -r requirements.txt
playwright install chromium
```

---

## Usage

### 1. Scan your existing spreadsheets (v2 with country + following)

**Sheet 1 (700-1000 usernames):**
```bash
python twitter_scanner_v2.py your_file.xlsx --sheet "700-1000" --output results_sheet1.xlsx --headless
```

**Sheet 2 (>1000 usernames) — verified only (faster):**
```bash
python verified_only_scanner.py your_file.xlsx --sheet ">1000" --output results_sheet2.xlsx --headless
```

**Sheet 2 with full enrichment (slower but more data):**
```bash
python twitter_scanner_v2.py your_file.xlsx --sheet ">1000" --output results_sheet2_full.xlsx --headless
```

**Custom follower threshold:**
```bash
python twitter_scanner_v2.py your_file.xlsx --min-followers 5000 --output results.xlsx
```

### 2. Scrape Kaito leaderboard + enrich with Twitter data

**Full pipeline (recommended):**
```bash
python kaito_pipeline.py --top 1000 --output kaito_creators.xlsx --headless
```

**Step by step (more control):**
```bash
# Step 1: Just scrape Kaito usernames first
python kaito_pipeline.py --top 1000 --kaito-only --output kaito_usernames.xlsx

# Step 2: Enrich with Twitter data later
python kaito_pipeline.py --enrich-from kaito_creators_kaito_checkpoint.json --output kaito_creators.xlsx --headless
```

**With Yaps scores (slower — rate limited to 100 API calls per 5 min):**
```bash
python kaito_pipeline.py --top 1000 --enrich-yaps --output kaito_creators.xlsx --headless
```

### 3. Scan both sheets at once
Edit `scan_both_sheets.py` configuration, then:
```bash
python scan_both_sheets.py
```

---

## Output Files

Each output `.xlsx` contains 3 sheets:

1. **Main Results** — Full data table with all columns, color-coded, auto-filtered
2. **Summary** — Quick stats (total, verified count, campaign-ready count, etc.)
3. **Campaign Ready** — Pre-filtered view sorted by follower count (highest first)

### Color Coding
- **Green** = Verified / Campaign Ready
- **Red** = Not Verified
- **Yellow** = Error / Not Found

---

## Estimated Times

| Task | Size | Estimated Time |
|------|------|----------------|
| Scanner v2 (full) | 100 usernames | ~10-15 min |
| Scanner v2 (full) | 700 | ~60-90 min |
| Scanner v2 (full) | 1000 | ~90-120 min |
| Verified-only | 1000 | ~50-70 min |
| Kaito scrape only (Step 1) | Top 1000 | ~5-10 min |
| Kaito + Twitter enrichment | Top 1000 | ~2-2.5 hours |

---

## Configuration

Tune these at the top of each script:

```python
MIN_DELAY = 3            # Min seconds between requests
MAX_DELAY = 6            # Max seconds between requests
BATCH_PAUSE_EVERY = 50   # Pause every N requests
BATCH_PAUSE_SECONDS = 30 # How long to pause (use 45 for 1000+)
CHECKPOINT_EVERY = 25    # Save progress every N usernames
MAX_RETRIES = 2          # Retry failed lookups
DEFAULT_MIN_FOLLOWERS = 1000  # Campaign-ready threshold
```

---

## Checkpoint / Resume

All scripts auto-save progress to `*_checkpoint.json` files. If a run stops:
- Just re-run the exact same command — it picks up where it left off
- Use `--no-resume` to start fresh

---

## Troubleshooting

### Twitter login wall
Run without `--headless`, manually log in when prompted, then let the script continue.

### Kaito leaderboard doesn't load all 1000
The leaderboard is JS-rendered. The scraper scrolls aggressively but may get fewer than 1000 if Kaito caps visible entries. Re-run or supplement with the Kaito API.

### "N/A" for country/region
Many users don't fill in their location. Typically 40-60% of profiles will have location data.

### Rate limiting / blocks
- Increase `BATCH_PAUSE_SECONDS` to 60
- Split into runs of 500
- Wait 15-30 min between batches
