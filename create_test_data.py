#!/usr/bin/env python3
"""
Simple test script to create test metrics data locally
No Google Sheets API needed - just for testing the dashboard
"""

import json
from pathlib import Path
from datetime import datetime

DATA_DIR = Path(__file__).parent / "data"
METRICS_FILE = DATA_DIR / "metrics.json"

# Create data directory
DATA_DIR.mkdir(exist_ok=True)

# Create test metrics
test_data = {
    "timestamp": datetime.now().isoformat(),
    "metrics": [
        {
            "name": "Portfolio Occupancy",
            "value": "87.50%",
            "updated": "2026-06-24",
            "source": "Occupancy Analysis"
        },
        {
            "name": "Leased Occupancy",
            "value": "92.30%",
            "updated": "2026-06-24",
            "source": "Occupancy Analysis"
        },
        {
            "name": "Percent Collected",
            "value": "94.20%",
            "updated": "2026-06-24",
            "source": "DQ Reports"
        },
        {
            "name": "Annualized Rent Growth",
            "value": "2.80%",
            "updated": "2026-06-24",
            "source": "Effective Rent Analysis"
        },
        {
            "name": "NOI Variance to Budget",
            "value": "TBD",
            "updated": "—",
            "source": "Monthly Financial Review"
        },
        {
            "name": "CapEx vs Budget",
            "value": "TBD",
            "updated": "—",
            "source": "Monthly Financial Review"
        },
        {
            "name": "Staffing Vacancies",
            "value": "TBD",
            "updated": "—",
            "source": "SharePoint Open Positions"
        }
    ],
    "count": 7,
    "success": True
}

# Write to file
with open(METRICS_FILE, 'w') as f:
    json.dump(test_data, f, indent=2)

print(f"✓ Test data created at: {METRICS_FILE}")
print(f"✓ Metrics: {test_data['count']}")
print(f"✓ Dashboard will load from: http://localhost:8000/data/metrics.json")
