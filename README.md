# DHF Generator for Wearable ECG Patch

This Python-based tool automates the creation of a Design History File (DHF) for a Wearable ECG Patch, a Class II medical device. By processing structured input files (CSV/TXT), it generates essential DHF documents aligned with FDA and ISO 13485 requirements.

---

## Features

- Automatically generates key DHF components from structured input files:
  - Design Overview
  - User Needs
  - Design Inputs & Outputs
  - Risk Management Table
  - Verification Methods
  - Change Log
- Integrates engineering and regulatory compliance documentation
- Saves structured `.txt` outputs for inclusion in submission-ready DHFs

---

## Input Files

Ensure the following files are in the same directory as `dhffinal.py`:

| File Name                 | Description                                      |
|--------------------------|--------------------------------------------------|
| `device_overview.txt`    | Summary of device purpose and components         |
| `user_needs.csv`         | Table of user needs and rationales               |
| `design_inputs.csv`      | Design inputs with linked user needs             |
| `design_outputs.csv`     | Outputs mapped to inputs                         |
| `risk_table.csv`         | Risk analysis (hazard, severity, mitigation)     |
| `verification_methods.csv` | Verification and validation strategies        |
| `change_log.csv`         | Log of design changes and rationales             |

---

## Requirements

- Python 3.6 or higher
- No external libraries required

---

## How to Run

From the command line:

```bash
python dhffinal.py
```

The script will read the input files and generate a series of documents summarizing the DHF content.

---
## Output Files

| File Name            | Description                                                  |
|---------------------|--------------------------------------------------------------|
| `DHF_Report.docx`    | A complete Word document covering all DHF sections           |
| `DesignControls.xlsx`| Structured Excel matrix mapping needs, inputs, outputs, tests |

---


## Device Overview

**Device**: Wearable ECG Patch  
**Class**: FDA Class II  
**Use Environment**: Out-of-hospital (remote monitoring)  
**End Users**: Cardiac patients and clinicians  

**Key Features**:
- Electrodes and analog front end
- Microcontroller and power supply
- Bluetooth Low Energy communication
- Posture recognition
- Software with mobile and PC interfaces

---
## DHF Report Content Overview

The generated DHF includes:

- **Device Overview**  
  Wearable ECG Patch with key features: wireless ECG sensing, BLE 5.0, skin-compatible adhesive, posture recognition, mobile & PC interface.

- **User Needs**  
  Comfort, mobility, accuracy, remote monitoring, long-term wear, ease of setup.

- **Design Inputs & Outputs**  
  Technical and functional requirements mapped with implementation evidence.

- **Verification Methods**  
  Bench tests, usability tests, wear trials, wireless performance testing.

- **Risk Management Table**  
  FMEA-style analysis with recommended mitigation actions.

- **Change Log**  
  Tracks revisions, authorship, and rationale.

- **Traceability Matrix**  
  User needs → inputs → outputs → verification methods

---

## Disclaimer

This DHF Generator is for educational and prototyping use only. It does **not** substitute for an official Quality Management System (QMS) or formal regulatory submission documentation. Always verify with your organization’s quality and regulatory experts.

---

## Author

Built by Shivangi Sarkar 

---

