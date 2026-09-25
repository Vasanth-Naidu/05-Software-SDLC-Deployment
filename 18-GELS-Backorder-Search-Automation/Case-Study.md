# 🛠️ Project 18: GELS Backorder Search Automation — Dual-Mainframe Inventory Alignment & Order Placement Engine

## Executive Overview:

* **Enterprise Context:** GE Capital International Services (GECIS) — GE Medical Systems Europe (GEMS Europe Operations)
* **Role:** Lead Automation Software Engineer & SDLC Owner (Citizen Developer)
* **The GEMBA Walk Discovery:** Identified through a proactive GEMBA Walk (Six Sigma) on the GEMS Europe operations floor. Conceptualised the problem statement, business case, ROI metrics, technical design, and project scope. Pitched and secured formal project sign-off from both GECIS India leadership and GEMS Europe Operations leadership.
* **Core Value Delivered:** Engineered a Visual Basic 5 (VB5) application that programmatically instantiated and controlled **two concurrent Hummingbird Host-Explorer terminal instances** using Object-Oriented Programming (OOPs) and terminal session objects. Automated daily cross-mainframe inventory searches, stock availability matching, replacement order placements, and end-of-day Excel reporting—executing, **03 FTEs worth of manual volume in under 03 hours** with zero human supervision during runtime.
* **Impact & Key Deliverables:**
  * **Capacity Compression (03 FTE's to 03 Hours):** Compressed a full 08-hour shift of manual data entry across 03 FTE's into an autonomous **03-hour background execution loop**, allowing the supervising representative to focus on high-value operational tasks.
  * **10,000 Parts Processed Daily:** Handled 9,000–10,000 daily line-item reviews across European hospital inventory queues with zero human omission errors, commission errors or keying mistakes.
  * **Zero-Data-Storage Architecture:** Engineered a privacy-compliant, zero-persistence pipeline (eliminating local database and storage) to strictly align with GEMS Europe data privacy mandates and General Data Protection Regulation (GDPR) regulations across European Economic Area (EEA) Countries, specifically addressing stringent Luxembourg standards.
  * **Latency-Adaptive Execution Engine:** Designed smart polling algorithms that dynamically paced keyboard macro injections based on Hummingbird terminal latency and mainframe response speeds, preventing dropped inputs or out-of-sync screen transitions.
  * **Catalysed Desktop Automation Strategy:** Earned personal commendations from GEMS Europe leadership for *"literally mimicking the keyboard and representative decision methodology,"* starting a floor-wide trend of converting high-volume, rule-based "low-hanging fruit" processes into citizen-led desktop automations.
* **Core Stack:** Visual Basic 5 (VB5 GUI & OOPs Terminal Session Objects), Hummingbird Host-Explorer (EHLLAPI/ OLE Automation Drivers), Dual-Instance Mainframe Terminals (3270/ 5250 Emulation), MS Excel API (Automated Report Generation).

---

## 1. Operational Challenge & Green-Screen Friction:

### Baseline Operational Friction:
GE Medical Systems (GEMS) Europe relied on Global Equipment Logistics System (GELS) mainframes to manage replacement parts for critical hospital diagnostic machinery (e.g., MRI coils, CT scanner tubes, X-ray detectors).

```text
┌─────────────────────────────────────────────────────────────────────────────────────────────┐
│                              BASELINE MANUAL WORKFLOW FRICTION                              │
├─────────────────────────────────────────────────────────────────────────────────────────────┤
│  • Volume & Human Dependency: 03 FTE's spent 08 hours daily manually reviewing 9,000 to     │
│    10,000 parts; absenteeism or leave immediately broke operational SLA delivery.           │
│  • Dual-Terminal Toggling: Representatives manually logged into Mainframe 'A' (GELS Queue), │
│    checked out-of-stock items, opened Mainframe 'B' (Central Stock), and keyed orders.      │
│  • High Fatigue & Error Risk: Repetitive green-screen keystrokes created fatigue, leading   │
│    to missed backorders and delayed spare parts delivery to field service engineers.        │
│  • Lack of Reporting: No standardised operational tracking existed for pending GELS         │
│    backorders, turnaround times (TAT), or regional demand breakdowns.                       │
└─────────────────────────────────────────────────────────────────────────────────────────────┘

```

* **GEMBA Walk Finding:** Uncovered during a process walk that 3 senior representatives were entirely absorbed in manual keystrokes, toggling screens, and copy-pasting part numbers across terminal windows.

---

## 2. System Architecture & Latency-Aware OOPs Loop:

The solution established a programmatic, dual-terminal object model in **VB5**:

```text
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                  GELS DUAL-MAINFRAME TERMINAL AUTOMATION PIPELINE                      │
├────────────────────────────────────────────────────────────────────────────────────────┤
│  ┌──────────────────────────┐                      ┌────────────────────────────────┐  │
│  │    GEMS Representative   │                      │      VB5 Runtime Engine        │  │
│  │   (Session Credentials)  │                      │   (Memory-Only Credentials)    │  │
│  └────────────┬─────────────┘                      └──────────────┬─────────────────┘  │
└───────────────┼───────────────────────────────────────────────────┼────────────────────┘
                │                                                   │
                ▼                                                   ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                        VB5 OBJECT-ORIENTED TERMINAL ORCHESTRATOR                       │
├────────────────────────────────────────────────────────────────────────────────────────┤
│  • Programmatically instantiates 2 Hummingbird Terminal Emulator Objects               │
│  • Mainframe Instance 'A': Scrapes European GELS out-of-stock backorder line items     │
│  • Mainframe Instance 'B': Verifies central inventory & triggers automated order fills │
│  • Adaptive Polling Engine: Dynamic latency sensing (waits for terminal screen buffer) │
└───────────────────────────────────────┬────────────────────────────────────────────────┘
                                        │
                                        ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                   END-OF-DAY AUTOMATED MS EXCEL REPORT & EXCEPTION LOG                 │
├────────────────────────────────────────────────────────────────────────────────────────┤
│ • Regional Summary   • Hourly Velocity   • User ID & TAT Metrics   • Pending Backorders│
└────────────────────────────────────────────────────────────────────────────────────────┘

```

### Key Architectural Pillars:

1. **Zero-Persistence Session Security:** The GEMS representative logged into the VB5 tool at shift start, entering mainframe credentials that were kept strictly in volatile memory (eliminated local database and storage to strictly align with GDPR regulations across EEA Countries).
2. **Programmatic Dual-Instance Control:** VB5 dynamically launched and took direct object control over **two parallel Hummingbird Host-Explorer terminal windows** via COM/OLE automation objects.
3. **Dual-Instance Matcher:**
   * **Instance 'A' (GELS Terminal):** Scraped out-of-stock backorder items and extracted part numbers, quantities, and regional identifiers.
   * **Instance 'B' (Central Stock Terminal):** Evaluated central stock availability and programmatically injected key buffers to execute replacement order purchases automatically.
4. **Latency-Aware Polling Loop:** Used adaptive delay loops that monitored screen buffer update flags in Hummingbird, dynamically slowing down or speeding up based on mainframe network latency to guarantee 100% input accuracy.
5. **In-Memory Error Logging & Reporting:** Trapped execution errors, screen mismatches, or out-of-stock exceptions directly in memory and compiled them into an automated end-of-day Excel workbook.

---

## 3. Operational Modules:

* **Module 1: Credential Gateway & OOPs Terminal Instantiator:** Accepts rep credentials, instantiates two Hummingbird OLE terminal session objects, and establishes secure connections to Mainframe Instance 'A' and Mainframe Instance 'B'.
* **Module 2: Latency-Adaptive Screen-Scraping & Key Injection Engine:** Scrapes line items from Mainframe Instance 'A' while dynamically monitoring screen readiness flags before firing keystrokes into Mainframe Instance 'B'.
* **Module 3: Exception Interceptor:** Captures system or order entry errors during processing, preventing terminal lock-ups and logging the exact failure reason for post-run review.
* **Module 4: End-of-Day MS Excel Reporting Engine:** Generates structured Excel workbooks detailing Region, Date, Hourly processing velocity, User ID, Turnaround Time (TAT), Pending GELS Backorders, and logged exception items.

---

## 4. Measurable Business Results & Operational Impact:

| Operational Dimension | 🛑 Baseline State (Pre-Automation) | 🎯 Post-Deployment State (GELS Engine) | ⚡ Operational Impact |
| --- | --- | --- | --- |
| **Execution Time** | 08 Hours (Full Shift across 03 FTE's) | **<03 Hours Autonomous Execution** | Compressed shift-long work into 03 hours, freeing reps for core tasks. |
| **Manual Effort & Capacity** | 03 Dedicated Full-Time Employees | **0 FTE Supervision Required** | Completely eliminated manual green-screen manipulation and human dependency. |
| **Daily Scale & Accuracy** | 9,000–10,000 parts reviewed manually daily | **10,000 Parts Processed Automatically** | Achieved 0% human omission errors, commission errors and eliminated keying mistakes. |
| **Data Privacy & Storage** | Manual screen reading | **100% Privacy Compliant** | Zero local database footprint; fully compliant with GEMS service agreements and GDPR regulations across EEA Countries. |
| **Process Continuity** | Absenteeism created severe backlogs | **100% Systemic Continuity** | The automation ran reliably regardless of team attendance or shift scheduling. |

---

## 5. Key Leadership & Technical Competencies Demonstrated:

* **GEMBA Walk Problem Discovery:** Walking the operational floor, identifying low-hanging operational bottlenecks, drafting formal business cases, and pitching solutions to global leadership.
* **Advanced Terminal Emulation & OOPs:** Instantiating and orchestrating concurrent legacy terminal objects via VB5 COM interface drivers.
* **Latency-Tolerant System Design:** Programming intelligent, state-aware polling logic that adapts automatically to back-end network and terminal response delays.
* **Strategic Desktop Automation Trendsetter:** Establishing a benchmark for citizen development at GECIS that proved small-scale, rule-based desktop automations deliver massive cumulative ROI.

---
