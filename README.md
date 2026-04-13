# ShiftGuard Parser API

This service converts nursing schedule Excel files into structured JSON.

## 🚀 What it does

- Accepts an uploaded Excel schedule
- Extracts all shift assignments
- Returns clean JSON in this format:

```json
[
  {
    "date": "2026-04-11",
    "name": "CLYDE SUNDARAM",
    "shift_type": "D",
    "description": "Day Shift",
    "status": "on_unit"
  }
]
