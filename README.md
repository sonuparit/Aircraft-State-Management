# Army Aviation Maintenance Automation (Excel VBA)

![alt text](MCC.png)

## Overview

During my service as an **Aviation Technician in the Indian Army**, aircraft maintenance data was managed through complex Excel sheets.
Many operational checks, maintenance schedules, and component tracking tasks were performed **manually**, which was time‑consuming and error‑prone.

To improve efficiency, I developed **VBA automation inside Excel** to streamline these workflows.

This automation significantly reduced repetitive work, improved accuracy, and helped technical teams quickly monitor aircraft maintenance status.

---

# Problem

Aircraft maintenance tracking required:

- Manual copying of aircraft state data between sheets
- Updating daily aircraft logs
- Calculating serviceable days
- Generating shareable aircraft state reports
- Monitoring components approaching maintenance limits

These tasks were performed **manually across multiple sheets**, which could take **hours or even weeks of effort** and carried a high risk of human error.

---

# Solution

I developed **five VBA automation modules** that transformed the workbook into an automated maintenance management tool.

Each module performs a specific operational function.

---

# Modules

## Module 1 — Aircraft State Initialization

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

## Module 2 — Daily Aircraft Log Entry

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

## Module 3 — Aircraft Serviceable Days Tracking

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

## Module 4 — Automated Aircraft State Export

**Purpose**

Automates the generation of a secure, formula-free export of the aircraft state report for daily transmission to base operations.

To meet operational and data-handling requirements, the report had to be shared in a format containing **values only**, without exposing formulas or internal workbook logic. This module converts the report into a sanitized copy and automatically saves it in a structured directory format, enabling fast and secure transfer over the local intranet.

**Key Automation**

- Creates structured folders automatically by **Year / Month**
- Generates a clean Excel copy containing only values
- Removes unnecessary data sections
- Saves the report automatically

**Impact**

- Standardizes reporting workflow
- Creates clean shareable maintenance reports

---

## Module 5 — Component Due Monitoring (Critical Automation)

**Purpose**

Monitors aircraft components approaching maintenance limits.

**Key Automation**

- Scans multiple aircraft sheets automatically
- Detects components with **less than 30 hours remaining**
- Detects components with **less than 30 days remaining**
- Consolidates all warnings into a monitoring sheet

**Impact**

This module was the **largest productivity improvement**.

Previously, technicians had to manually inspect hundreds of rows across multiple aircraft sheets.

This automation:

- Scans all aircraft automatically
- Identifies critical components instantly
- Consolidates warnings into a single sheet

**Manual process:** ~2 weeks of effort  
**Automated process:** a few seconds

This dramatically improved **maintenance awareness and safety monitoring**.

---

# Technologies Used

- Microsoft Excel
- VBA (Visual Basic for Applications)
- Spreadsheet automation
- File system automation

---

# Skills Demonstrated

- Process automation
- Operational workflow optimization
- Data processing automation
- Technical problem solving
- Real‑world system efficiency improvements

---

# Security Notice

The original workbook contains **sensitive operational data**, therefore the file is locked and not publicly shared.

A **live demonstration can be presented upon request on my system.**
