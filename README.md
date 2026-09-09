# CS Pulse

Streamlit dashboard for the GC Customer Support team -- Chats, Calls and Adherence,
company-wide and per agent. Sibling to Ops Pulse (same dark theme, same metric-card style).

## Run it

```bash
pip install -r requirements.txt
streamlit run app.py
```

## What you upload

One workbook with 5 tabs, matching the live source sheet:

- **Agents ID** -- Agent Name, Agent ID, (optional) Team
- **Schedule** -- long format: Employee name, Date, Shift, Is WFH
- **Calls** -- raw Calls export
- **Chats** -- raw Chats export (Assignee is the numeric Agent ID)
- **Agents Activity** -- raw presence log (Agent Name, Timestamp, State)

Pick a date range in the sidebar (defaults to Jul 1 - Aug 31 2026, the two closed months
used to validate this build) and the report recomputes.

## Rules baked into `logic.py` (confirmed with Mahmoud, Sep 2026)

- Bahrain team (Sayed Hadi Alwedaei, Fatima Hassan, Zainab Abbas, Mahdi Ali) excluded entirely.
- Logistics-team agents (Heba Tarek, Mayar Khaled, Nada Esaam) have no shifts -- Adherence
  stays blank for them, Chats/Calls still counted.
- Shift planned minutes: 9 AM-5 PM / 11 AM-7 PM / 4 PM-12 AM = 420 min (7h, 1h unpaid break
  baked in); 12 AM-9 AM (night shift; "11:59 PM - 9 AM" and "12 PM - 9 AM" are the same
  shift, both normalized to this) = 480 min.
- Day Off / Ann / CL / Sick / PH / Termination = non-working, excluded from Adherence. An
  agent is dropped from Adherence entirely from their first Termination date onward.
- "Task" / "Task - Normal Shift" -- meaning still unconfirmed. Excluded from Adherence and
  surfaced separately in the Diagnostics section rather than guessed at.
- Company policy: female agents (gender inferred from first name) on the 4 PM-12 AM shift
  are always treated as work-from-home -- 9-hour shift ending 1 AM -- regardless of the
  Is WFH flag that day.
- Calls State is always shown as all 8 individual states (Serviced/Successful/Dropped/No
  Answer/Abandoned/Blocked/Busy/Failed), never collapsed into one Answered/Missed number,
  so the underlying cause stays visible per agent. Answered = Serviced + Successful.
- FCR = no return contact from the same customer (Contact ID) within 7 days of a resolved
  chat's close time. Computed only within the uploaded/filtered Chats data -- a contact's
  next conversation just past a period boundary can't be seen.
- Occupancy % is a calls-only metric (Busy / (Busy+Online) time) -- chat handling shows as
  "Available" in the presence log, same as idle time, so it reads low by design.

## Known open items (see Diagnostics tab in the app)

- Nariman Shedid vs "Nariman Ezzat" and Samaa Ahmed vs "Samaa Aabullwahab" surname splits --
  still merged as one canonical person each, unresolved with HR.
- Schedule's "Hagar Shaban" is kept as a SEPARATE person from Agents ID's "Hagar Ahmed" --
  believed to be two different people, not a spelling variant.
- Mohamed Bassem (real Calls/Activity signal) is not on the Agents ID roster or Schedule --
  excluded from this build until added.
- Karim Mohamed, Mostafa Gomaa, Abdallah Mahmoud, Amr Hazem, Salma Adel, Duha Younis --
  weak or zero Chats/Calls/Activity signal; will simply not appear in the per-agent tables.
