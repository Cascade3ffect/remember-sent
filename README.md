# remember-sent

Outlook conversation analysis and dashboard for tracking commitments and follow-ups.

The application analyzes Sent Items and another received mail folder from Microsoft Outlook, identifies potential commitments, classifies conversations as OPEN, WAITING, or CLOSED, and provides a web-based dashboard backed by SQLite.

---

## Features

- Outlook COM integration
- Conversation-level email analysis
- Commitment detection using configurable regex patterns
- Identification of reply status
- Status tracking:
  - OPEN
  - WAITING
  - CLOSED
- Priority scoring using ActionScore
- CSV export with timestamped results
- SQLite database storage
- Browser-based dashboard using aiohttp and Jinja2
- Historical import tracking

---

## Requirements

### Software

- Microsoft Outlook (Desktop)
- Python 3.13+
- PowerShell 5.1 or newer

### Python Packages

```powershell 
pip install aiohttp aiohttp-jinja2 aiosqlite jinja2 pandas
```

---

## Directory Structure

```text
EmailDashboard/
│
├── Execute_v01.bat
├── app.py
├── import_csv.py
├── email.db
│
├── .venv/
│
├── imports/
│   ├── SentEmailAnalysis_Hybrid_v07.ps1
│   └── *.csv
│
├── templates/
│   └──dashboard.html
│
├── static/
│   └── site.css
│
└── README.md
```

# Setup

### Clone Repository

```Powershell
git clone https://github.com/Cascade3ffect/remember-sent.git
cd EmailDashboard
```

### Create Virtual Environment

```powershell
py -3 -m venv .venv
```

### Activate Virtual Environment

```powershell
.venv\Scripts\Activate.ps1
```

### Install Dependencies *** Feature Not Tested *** 

```powershell
pip install -r requirements.txt
```

## Running the Application

### Recommended Method

Simple double click:

```text
Execute_v01.bat
```

The batch file performs the complete workflow automatically.


### What Execute_v01.bat Does

### 1. Temporarily Allow PowerShell Script Execution

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

This change only applies to the current PowerShell session.


### 2. Import Latest Emails from Outlook

Runs:

```powershell
.\imports\SentEmailAnalysis_Hybrid_v07.ps1
```

This script:

- Requests Outlook mailbox, folder name, and timeframe from user
- Reads Outlook mail
- Analyzes conversations
- Calculates commitment status
- Creates a timestamped CSV export

Output example:

```text
imports\
├── SentEmailAnalysis_Hybrid_v07_results_20261123_1535.csv
```


### 3. Activate Python Environment

Runs:

```powershell
.venv\Scripts\Activate.ps1
```

This ensures the required Python packages are available.

---

### 4. Convert Latest CSV to SQLite Database

Runs:

```powershell
python import_csv.py
```

This step:

- Finds the newest CSV file
- Imports results into SQLite
- Updates conversation records
- Records import history

Database:

```text
email.db
```

Tables:

```text
conversations
imports
```

### 5. Launch Dashboard

Runs:

```powershell
python app.py
```

The web application starts locally at:
```text
http://localhost:8080
```

## Dashboard Columns

| Column | Description |
|----------|----------|
| DaysSinceLastMessage | Days since latest activity |
| OutstandingCommitments | Open commitments detected |
| Subject | Email subject |
| LastMessageDate | Timestamp of latest activity |
| LastDirection | Sent or Received |
| CommitmentCount | Commitments detected |
| ClosedCommitmentCount | Commitments considered completed |
| ActionScores| Priority ranking |
| ConversationID | Outlook ConversationID |
| Status | OPEN, WAITING, or CLOSED |
| Body | Truncated message body |


## Status Definitions

### OPEN
Conversation contains an outstanding commitment and the most recent message was sent by you.

Meaning:
```text
Action required by you.
```

### WAITING

Conversation contains an outstanding commitment but the most recent message was received from someone else.

Meaning:

```text
Waiting for somebody else.
```


### CLOSED

No outstanding commitments remain.

Meaning:

```text
No further action required.
```


## Database Schema

### conversations

Contains the most recent dashboard snapshot.

### imports *** Feature in Process ***
Stores import history.

```sql
CREATE TABLE imports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    imported_at TEXT,
    source_file TEXT);
```

This enables future reporting such as:

- Open commitments over time
- New conversations
- Closed conversations
- Trend analysis

## Troubleshooting

### Outlook Script Execution Disabled

Run:
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```


### Python Not Found
Verify:

```powershell
.venv\Scricts\python.exe --version
```

---

### No Dashboard Data

Verify database contains records:

```powershell
sqlite3 email.db
select count() from conversations;
```

---

### Server Starts But Browser Does Not Open

Navigate manually:

```text
http://localhost:8080
```

---

## Future Enhancements

- Commitment history tracking - Trend charts
- Search and filtering
- Dashboard KPI cards
- Project-level reporting
- Email link-back to Outlook conversation 
- Automatic scheduled imports via Windows Task Scheduler

---

## License

Internal Use
