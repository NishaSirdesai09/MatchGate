# MatchGate

Alumni match review dashboard with a Google Apps Script backend.

## Pieces

| Piece | File | Role |
|-------|------|------|
| Static prototype | `matchGate.html` | localStorage-only demo |
| Backend | `code.gs` | Shared reviews + records via Google Sheet |

## Backend

Deploy `code.gs` as an Apps Script Web App (Execute as Me, access Anyone).
Sheet tabs: `Sheet1` (reviews), `Records` (people).
