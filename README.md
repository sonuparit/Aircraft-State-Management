# 🛩️ Army Aviation Maintenance Automation (Excel VBA)

### 🏅 Received a **commendation from the Officer Commanding** for developing this automation system, which improved operational efficiency during high-pressure (**`Pulwama attack`**) conditions.

These automations are now **`actively used across aviation units`**, and I’ve received hundreds of calls — both for support and in recognition of the impact they’ve created.

**Actual system snapshots at bottom**

![alt text](MCC.png)

## 📚 Table of Contents

- [Overview](#-overview)
- [Problem](#️-problem)
- [Solution](#-solution)
- [Modules](#-modules)
  - [Module 1 — Aircraft State Initialization](#module-1--aircraft-state-initialization)
  - [Module 2 — Daily Aircraft Log Entry](#module-2--daily-aircraft-log-entry)
  - [Module 3 — Serviceable Days Tracking](#module-3--aircraft-serviceable-days-tracking)
  - [Module 4 — Automated Aircraft State Export](#module-4--automated-aircraft-state-export)
  - [Module 5 — Component Due Monitoring](#module-5--component-due-monitoring-critical-automation)
- [System Snapshots](#-system-snapshots-actual-implementation)
- [Technologies Used](#-technologies-used)
- [Skills Demonstrated](#-skills-demonstrated)
- [Security Notice](#-security-notice)

## 📖 Overview

During my service as an **Aviation Technician in the Indian Army**, aircraft maintenance data was managed through complex Excel sheets.\

Many operational checks, maintenance schedules, and component tracking tasks were performed **manually**, which was time‑consuming and error‑prone.

To solve this, I developed a **VBA-powered automation system inside Excel** that transformed the workbook into an efficient maintenance management tool.

This solution significantly reduced manual effort, improved accuracy, and enabled faster operational decision-making.

---

## ⚠️ Problem

Aircraft maintenance tracking required:

- Manually copying data across multiple sheets
- Updating daily aircraft logs
- Calculating serviceable days
- Generating shareable reports
- Monitoring components nearing maintenance limits
- Hard to find old records

These tasks were performed **manually across multiple sheets**, which could take **hours or even weeks of effort** and carried a high risk of human error.

---

## 💡 Solution

I developed **five VBA automation modules** that transformed the workbook into an automated maintenance management tool.

Each module performs a specific operational function.

---

## 🧩 Modules

### Module 1 — Aircraft State Initialization

**Purpose**

Prepares and resets the aircraft state sheet for daily updates.

**Key Automation**

- Copies aircraft state data to backup reference sections
- Resets daily time counters
- Initializes the worksheet for new operational entries

**Impact**

- Eliminates repetitive manual setup steps
- Ensures consistent daily data initialization

---

### Module 2 — Daily Aircraft Log Entry

**Purpose**

Automates insertion of daily aircraft log entries into maintenance records.

**Key Automation**

- Extracts aircraft operational data from the main sheet
- Inserts new entries into aircraft log sheets
- Maintains historical records automatically

**Impact**

- Prevents manual data entry errors
- Speeds up aircraft log updates

---

### Module 3 — Aircraft Serviceable Days Tracking

**Purpose**

Tracks aircraft availability and serviceable days.

**Key Automation**

- Inserts daily status entries
- Captures aircraft operational data
- Updates serviceability tracking sheets

**Impact**

- Provides a quick overview of aircraft availability
- Simplifies operational readiness tracking

---

### Module 4 — Automated Aircraft State Export

**Purpose**

Automates the generation of a secure, formula-free export of the aircraft state report for daily transmission to base operations.

To meet operational and data-handling requirements, the report had to be shared in a **values-only format**, without exposing formulas or internal workbook logic. This allowed the base to directly import the data into their systems for further processing.

This module converts the report into a sanitized copy and automatically saves it in a structured directory format, enabling fast and secure transfer over the local intranet.

**Key Automation**

- Creates structured folders automatically by **Year / Month**
- Generates a clean Excel copy containing only values
- Removes unnecessary data sections
- Saves the report automatically

**Impact**

- Standardizes reporting workflow
- Creates clean shareable maintenance reports
- Instantly became **digital archive of historical reports**
- Eliminated reliance on years of printed records
- Enables **instant retrieval of past data**

---

### Module 5 — Component Due Monitoring (Critical Automation)

**Purpose**

Monitors aircraft components approaching maintenance limits.

**Key Automation**

- Scans all aircraft sheets automatically
- Flags components with:
  - Less than **30 hours remaining**
  - Less than **30 days remaining**
- Consolidates alerts into a single dashboard

**Impact**

This module was the **largest productivity improvement**.

Previously, technicians had to manually inspect hundreds of rows across multiple aircraft sheets.

This automation was born out of necessity during the Pulwama attack, when operations were running day and night to keep aircraft mission-ready. I created it to reduce manual effort and find brief moments of rest amid intense operational pressure.

> 🏅 Received a **commendation from the Officer Commanding** for developing this automation system, which improved operational efficiency during high-pressure conditions.

**Manual process:** ~2 weeks of effort  
**Automated process:** a few seconds

This dramatically improved **maintenance awareness and safety monitoring**.

---

## 🧰 Technologies Used

- Microsoft Excel
- VBA (Visual Basic for Applications)
- Spreadsheet automation
- File system automation

---

## 🧠 Skills Demonstrated

- Process automation
- Operational workflow optimization
- Data processing automation
- Technical problem solving
- Real‑world system efficiency improvements

---

## 📸 System Snapshots (Actual Implementation)

Below are real screenshots from the automation system used during operations.

These demonstrate the scale, structure, and functionality of the solution.

`Data is from 2021 and screenshot taken on my system on 2026`

### 🎛️ Automation Control Panel

Centralized control interface for executing automation modules.

- Main AC state
- AC reset for daily initialization
- One-click execution of workflows
- Backup generation (exact values only copy) date wise
- Component due checks (30h / 30d)
![alt text](screenshots/screenshot02.png)

### ⏱️ Component Due Monitoring

Displays components approaching maintenance limits (hours/days).

- Highlights critical components automatically
- Consolidates data from multiple aircraft
- Eliminates manual inspection across sheets
![alt text](screenshots/screenshot01.png)

### 📊 Aircraft Status Tracking

Provides a quick overview of aircraft availability and operational status.

- Tracks multiple aircraft in a single view
- Simplifies daily monitoring
![alt text](screenshots/screenshot03.png)

### 📄 Individual AC Lifecycle Record

Detailed tracking of individual components.

- Tracks installation, usage, and remaining life
- Supports accurate maintenance planning
![alt text](screenshots/screenshot04.png)


---

## 🔒 Security Notice

The original workbook contains **sensitive operational data**, therefore the file is locked and not publicly shared.

A **live demonstration can be presented upon request on my system.**
