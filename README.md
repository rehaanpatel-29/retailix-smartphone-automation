
# Retailix Smartphone Launch Automation

An automated Python workflow for discovering smartphone launch information, verifying product details, and preparing structured product data for Retailix.

## Project Overview

This project automates parts of the smartphone product research and data preparation process. It combines Python-based processing with GitHub Actions to support scheduled execution and maintain a structured Excel product template.

## Key Features

- Automated execution through GitHub Actions.
- Smartphone launch discovery using configured sources.
- Product information verification against relevant sources.
- Structured product data preparation for Excel import.
- Duplicate prevention for existing product combinations.
- Scheduled workflow execution and manual workflow dispatch.

## Technology Stack

| Technology | Purpose |
| --- | --- |
| Python | Automation and data processing |
| GitHub Actions | Workflow scheduling and execution |
| Excel | Structured product data and import template |
| Requests and BeautifulSoup | Web requests and HTML parsing |
| Playwright | Browser-based page processing, where configured |

## Repository Structure

| File or Directory | Description |
| --- | --- |
| `main.py` | Main automation script |
| `.github/workflows/` | GitHub Actions workflow configuration |
| `product_import_template (1).xlsx` | Excel product import template |

## Workflow

1. Start the automation manually or through its configured schedule.
2. Discover potential smartphone launch announcements.
3. Validate qualifying product information using configured sources.
4. Process eligible product records and check for duplicates.
5. Prepare the structured product data in the Excel template.
6. Save changes according to the configured workflow.

## Execution

The automation is designed to run through GitHub Actions.

To execute it manually:

1. Open the repository's **Actions** tab.
2. Select the smartphone automation workflow.
3. Choose **Run workflow**, if enabled.

The workflow's schedule and execution settings are defined in the YAML file under `.github/workflows/`.

## Project Purpose

The objective is to reduce repetitive product research and data preparation while improving consistency and efficiency in the smartphone import workflow.

## Author

**Rehan Patel**

Computer Science Graduate | Data Analyst Intern

GitHub: https://github.com/rehaanpatel-29

---

*Project documentation describes the intended workflow. Actual source verification, duplicate handling, and file updates depend on the current implementation and configuration.*
