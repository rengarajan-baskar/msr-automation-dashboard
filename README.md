MSR Automation Dashboard

## Automated Monthly Status Review (MSR) Reporting & Analytics Platform

### Overview

The **MSR Automation Dashboard** is an end-to-end reporting and analytics solution designed to transform raw support ticket data into structured, management-ready insights. Built using **Python, Pandas, Streamlit, and Plotly**, this solution eliminates manual Excel-based reporting processes and replaces them with an intelligent, interactive, and scalable automation framework.

This project was developed to address a common operational inefficiency within support teams — the time-consuming and error-prone preparation of Monthly Status Review (MSR) reports. By automating data ingestion, classification, aggregation, visualization, and export workflows, the solution reduces preparation time from several hours to just a few minutes while significantly improving consistency and analytical depth.

---

## Business Problem

Support teams often rely on manually maintained Excel trackers for incident, service request, and problem management reporting. Preparing MSR reports typically involves:

* Repeated filtering and pivoting in Excel
* Manual categorization of ticket types and root causes
* Counting and summarizing incidents by priority, state, and assignment group
* Creating visual summaries for management review
* Exporting and formatting reports every month

This manual approach:

* Consumes significant operational time
* Introduces risk of human error
* Produces inconsistent reporting structures
* Limits analytical depth

The MSR Automation Dashboard directly solves these inefficiencies.

---

## Solution Architecture

The system is designed using a clean separation of concerns:

### Backend Processing Engine (`msr_automator.py`)

Responsible for:

* Reading Excel ticket trackers
* Normalizing inconsistent column structures
* Applying rule-based root cause inference
* Generating pivot summaries
* Calculating performance metrics (e.g., MTTR)
* Producing downloadable Excel reports

This module is reusable, scalable, and configuration-driven.

---

### Interactive Analytics Dashboard (`msr_app.py`)

Built with **Streamlit**, this module provides:

* Excel file upload capability
* Dynamic filtering by:

  * Ticket Type
  * Priority
  * State
  * Opened Date Range
  * Resolved Date Range
* Real-time summary metrics
* Interactive charts using Plotly
* Downloadable structured MSR reports

The dashboard enables non-technical users to generate reports without interacting with code.

---

### Configuration Layer (`config.yaml`)

A YAML-based configuration file allows:

* Flexible column alias mapping
* Customizable root cause keyword rules
* Output sheet structure control

This ensures adaptability without requiring code changes when tracker formats evolve.

---

## Key Features

* Automated ticket categorization (Incident, SCTask, Problem, etc.)
* Rule-based root cause classification
* Priority and state distribution analysis
* Assignment group and customer segmentation
* SLA breach tracking
* MTTR (Mean Time to Resolve) calculation
* Monthly trend analysis
* Interactive bar, pie, and line charts
* Filter-driven dynamic dashboards
* One-click downloadable Excel summaries
* Config-driven extensibility
* Clean modular architecture

---

## Technology Stack

* **Python 3.x**
* **Pandas** – Data processing and aggregation
* **Streamlit** – Web interface
* **Plotly** – Interactive data visualization
* **PyYAML** – Configuration management
* **XlsxWriter** – Excel export generation

---

## Business Impact

* Reduces MSR preparation time from hours to minutes
* Eliminates repetitive manual filtering
* Improves reporting consistency and standardization
* Enables data-driven decision making
* Scalable for larger ticket volumes
* Demonstrates automation and analytics capability

---

## Security & Deployment

* Designed for local execution (localhost)
* Can be deployed internally within enterprise networks
* Does not transmit data externally
* Suitable for controlled operational environments

---

## Summary

> The MSR Automation Dashboard is a Python-based analytics platform that transforms raw support ticket Excel data into structured, interactive, and exportable management reports. It replaces manual MSR preparation with a scalable, automated, and data-driven reporting framework, significantly improving operational efficiency and analytical visibility.
