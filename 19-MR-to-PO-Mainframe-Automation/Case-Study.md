# 🛠️ Project 19: MR to PO Mainframe Automation — Automated Purchase Order Execution Engine

## Executive Overview:

* **Enterprise Context:** GE Capital International Services (GECIS) — GE Medical Systems USA (GEMS USA Operations/ Global Asset Management - GAM India)
* **Role:** Lead Automation Software Engineer & SDLC Owner (Citizen Developer)
* **Best Practice Replication Origin:** Spurred directly as a cross-account best-practice expansion from the success of **Project 18 (GELS Backorder Automation)**. Recognised by GEMS USA leadership for pioneering work on mainframe emulator terminal automations, GEMS USA leadership invited me to solve a similar manual Purchase Order entry bottleneck within their Global Asset Management (GAM) India team.
* **Core Value Delivered:** Concepted, engineered, deployed, and maintained **MR to PO Mainframe Automation**, a Visual Basic 5 (VB5) desktop application integrated with Hummingbird Host-Explorer terminal emulation. The engine automated the ingestion of daily FTP Material Request (MR) reports, cross-referenced exceptional vendor/part verification rules, auto-calculated pricing entries, and executed direct Purchase Order (PO) keying into the legacy USA mainframe—with zero local database persistence, pushing all completion outputs directly to MS Excel.
* **Impact & Key Deliverables:**
  * **$67,690 Hard Cost Savings:** Delivered direct, auditable financial savings by slashing manual processing cycle times and eliminating high-cost human keying errors.
  * **80%+ Cycle Time Reduction per PO:** Compressed processing time from **05–06 minutes per PO down to seconds**, drastically accelerating part fulfilment to GE Warehouses across North America.
  * **Zero-Data-Storage Architecture:** Engineered a privacy-compliant, zero-persistence pipeline (eliminating local database storage) that loads FTP data in memory and writes output directly to MS Excel workbooks to strictly align with GEMS data privacy mandates.
  * **Unified Single-Screen Decision & Keying Interface:** Replaced manual toggling between spreadsheets, vendor lookup sheets, and green-screen terminals with an integrated VB5 GUI that presented decision-making context and executed order placements in a single click.
  * **Embedded Exception & Typo Interceptor:** Built strict programmatic validation rules that pre-screened material requests for exceptional vendors/parts and validated unit prices before committing transactions to the mainframe.
* **Core Stack:** Visual Basic 5 (VB5 GUI & OOPs Terminal Session Objects), Hummingbird Host-Explorer (EHLLAPI/ OLE Automation Drivers), Mainframe Terminal Sessions (3270 Emulation), FTP File Parser (In-Memory Report Ingestion), MS Excel API.

---

## 1. Operational Challenge & Green-Screen Keying Bottlenecks:

### Baseline Operational Friction:
The Global Asset Management (GAM) India team within GEMS USA managed proactive inventory replenishment across GE Warehouses across North America. Frontline representatives manually converted Material Requests (MR) into official Purchase Orders (POs) inside the legacy mainframe:

```text
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                              BASELINE MANUAL WORKFLOW FRICTION                         │
├────────────────────────────────────────────────────────────────────────────────────────┤
│  • FTP Ingestion & Spreadsheet Toggling: Representatives manually downloaded daily FTP │
│    reports containing raw MR lines, parsing hundreds of part orders across MS Excel.   │
│  • High Touch Time (05–06 Mins/ PO): Each order required manually checking exceptional │
│    vendor codes, part rules, unit prices, and then keying values into 3270 screens.    │
│  • High Error Risk: Manual pricing entry led to costly typo errors, while skipped      │
│    vendor validations caused misrouted orders and warehouse fulfilment delays.         │
│  • Double-Data Entry: Reps had to key orders into the mainframe AND manually update    │
│    Excel tracking logs post-placement, doubling non-value-added (NVA) effort.          │
└────────────────────────────────────────────────────────────────────────────────────────┘

```

* **Best Practice Opportunity:** Following the floor-wide success of Project 18, GEMS USA leadership invited me to replicate the terminal emulation architecture to solve this manual keying bottleneck.

---

## 2. System Architecture & Automated Ingestion Loop:
The solution established an automated FTP-to-Mainframe pipeline powered by **VB5** and **Hummingbird Terminal Emulation**:

```text
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                      MR TO PO AUTOMATED EXECUTION PIPELINE                             │
├────────────────────────────────────────────────────────────────────────────────────────┤
│  ┌──────────────────────────┐                      ┌────────────────────────────────┐  │
│  │      Daily FTP Server    │                      │ Automated FTP File Ingestion   │  │
│  │   (Raw MR Data Reports)  │                      │ & Memory-Only Parser (VB5)     │  │
│  └────────────┬─────────────┘                      └──────────────┬─────────────────┘  │
└───────────────┼───────────────────────────────────────────────────┼────────────────────┘
                │                                                   │
                ▼                                                   ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                       VB5 SINGLE-SCREEN DECISION & VALIDATION GUI                      │
├────────────────────────────────────────────────────────────────────────────────────────┤
│  • Automated Exceptional Vendor & Part Verification Lookup                             │
│  • Auto-Price Calculation & Typo Prevention Engine                                     │
│  • One-Click/ Autonomous Mainframe PO Placement (Hummingbird OLE Session Object)       │
└───────────────────────────────────────┬────────────────────────────────────────────────┘
                                        │
                                        ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│             DIRECT MS EXCEL OUTPUT & TRACKING LOG (ZERO LOCAL STORAGE)                 │
├────────────────────────────────────────────────────────────────────────────────────────┤
│ • Successful PO Confirmation   • Pricing Audit Log   • Exception Vendor Routing Queue  │
└────────────────────────────────────────────────────────────────────────────────────────┘

```

### Key Architectural Pillars:

1. **Automated FTP Ingestion & Parsing:** Programmatically connected to GEMS FTP servers, pulled daily MR reports, and loaded pending parts directly into the volatile VB5 memory buffer (no disk persistence).
2. **Embedded Rule Engine & Exception Interceptor:** Pre-screened each line item against exceptional vendor directories, restricted part lists, and pricing rules—flagging anomalies *before* mainframe entry to eliminate human error.
3. **Single-Screen Unified GUI:** Displayed all decision-making context (part specs, vendor history, calculated unit cost) alongside an instant "Execute PO" interface, eliminating multi-window toggling.
4. **Hummingbird Key Buffer Injector:** Used OOPs terminal session drivers to programmatically fire keystrokes, populate screen fields, and navigate 3270 mainframe menus with zero manual typing.
5. **Direct-to-Excel Log Synchronisation:** Streams fulfilment logs and PO updates directly into MS Excel workbooks in real time without writing to a local database.

---

## 3. Operational Modules:

* **Module 1: FTP Ingester & Memory Parser:** Automatically retrieves daily MR data files from GEMS FTP servers and parses raw text lines into structured volatile memory objects.
* **Module 2: Exceptional Vendor & Price Validation Engine:** Cross-references part numbers against approved vendor matrices and validates unit pricing to prevent keying typos.
* **Module 3: Unified Single-Screen Decision Workspace:** Presents representatives with a clean GUI containing all necessary decision metrics and one-click PO execution controls.
* **Module 4: Terminal Injection & Direct-to-Excel Sync Module:** Fires automated keystrokes into the Hummingbird 3270 terminal instance and writes completed PO numbers directly to MS Excel tracking logs.

---

## 4. Measurable Business Results & Operational Impact:

| Operational Dimension | 🛑 Baseline State (Pre-Automation) | 🎯 Post-Deployment State (MR-PO Engine) | ⚡ Operational Impact |
| --- | --- | --- | --- |
| **Direct Financial ROI** | High manual error costs & NVA labour | **$67,690 Hard Savings Delivered** | Direct, auditable bottom-line financial impact. |
| **Processing Time/ PO** | 05–06 Minutes per PO (Manual Entry) | **Sub-Minute/ Automated Execution** | **80%+ cycle time reduction** per purchase order. |
| **Data Quality & Accuracy** | Frequent typos in prices & vendor codes | **100% Rule Validation & Precision** | Eliminated keying typos and exceptional vendor errors. |
| **Data Privacy & Storage** | Manual spreadsheet saving | **Zero Local Database Footprint** | Pushed output directly to MS Excel to maintain zero data persistence. |
| **Process Complexity** | Toggling FTP, Excel, Binders & Mainframe | **Unified Single-Screen Workspace** | Integrated all decision data and execution into one GUI. |

---

## 5. Key Leadership & Technical Competencies Demonstrated:

* **Best Practice Replication & Cross-Pollination:** Successfully adapting and scaling a proven terminal automation architecture (Project 18) to solve distinct operational challenges in new business units (GEMS USA).
* **Hard Cost Reduction Engineering:** Aligning technical software design directly with measurable financial outcomes ($67,690 hard savings).
* **Privacy-Compliant Desktop Architecture:** Ensuring zero local database footprint by streaming runtime data through volatile memory directly into MS Excel outputs.
* **Poka-Yoke (Mistake-Proofing) System Design:** Inserting automated validation controls into desktop software to prevent human typo errors before they hit core mainframes.


---
