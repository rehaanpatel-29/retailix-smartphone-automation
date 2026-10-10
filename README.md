
# Retailix Smartphone Product Automation

**Automating Product Data Preparation to Streamline Retailix Product Management**

## Overview

The Retailix Smartphone Product Automation project was developed to address operational challenges associated with manually adding smartphone products to the Retailix platform.

Manual product entry requires repetitive data collection, verification, formatting, and preparation. This process can consume valuable time and introduce inconsistencies into product records.

To address these challenges, this project introduces a Python-based automation workflow that streamlines smartphone launch discovery and product data preparation, helping reduce repetitive manual effort and improve data consistency.

## Business Problem

Adding smartphone products to Retailix manually presented several challenges:

- Repetitive collection of product information from multiple sources.
- Time-consuming verification of smartphone launch details and specifications.
- Manual preparation of product records in the required import format.
- Risk of duplicate entries and inconsistent product information.
- Repeated effort required to maintain up-to-date product data.

These challenges created an opportunity to improve the product onboarding workflow through automation.

## Proposed Solution

The project uses Python and GitHub Actions to automate configured stages of smartphone product research and data preparation.

The workflow is designed to identify potential smartphone launches, verify relevant product information against configured sources, process qualifying records, check for existing product combinations, and prepare structured data for the Retailix product import template.

The objective is to simplify the workflow while maintaining a consistent and structured approach to product data management.

## Key Features

- **Automated Discovery:** Identifies potential smartphone launch information using configured discovery sources.
- **Product Verification:** Validates relevant product details against configured information sources.
- **Structured Data Preparation:** Organizes product records according to the Excel import template.
- **Duplicate Prevention:** Checks existing product combinations before processing eligible records.
- **Scheduled Execution:** Uses GitHub Actions for configured automated runs and manual workflow execution.
- **Repeatable Workflow:** Reduces reliance on repetitive manual research and data preparation.

## Technology Stack

| Technology | Application |
| --- | --- |
| Python | Automation, data processing, and workflow logic |
| GitHub Actions | Scheduled execution and workflow orchestration |
| Excel | Structured product data and import preparation |
| Requests | HTTP requests to configured sources |
| BeautifulSoup | HTML parsing and content extraction |
| Playwright | Browser automation where required by the workflow |

## Workflow Architecture

1. **Discovery:** Identify potential smartphone launch announcements.
2. **Verification:** Check relevant launch information and product details against configured sources.
3. **Data Processing:** Extract and organize eligible product information.
4. **Duplicate Checking:** Compare product combinations against existing records.
5. **Data Preparation:** Format qualifying records for the Retailix Excel import template.
6. **Execution and Maintenance:** Run the workflow through GitHub Actions and maintain the project as requirements evolve.

## Repository Structure

| File or Directory | Description |
| --- | --- |
| `main.py` | Main Python automation script |
| `.github/workflows/` | GitHub Actions workflow configuration |
| `product_import_template (1).xlsx` | Product import template used by the workflow |

## Business Value

The project is intended to improve the efficiency of smartphone product onboarding by:

- Reducing repetitive manual product research and data entry.
- Standardizing product data preparation.
- Helping identify duplicate product combinations.
- Supporting a repeatable product data workflow.
- Making routine product updates easier to manage.

Actual time savings and accuracy improvements should be measured against the previous manual process.

## Execution

The automation can be triggered through the configured GitHub Actions workflow.

1. Open the repository's **Actions** tab.
2. Select the smartphone automation workflow.
3. Use **Run workflow** when manual execution is enabled.

Scheduled execution depends on the workflow configuration. Python dependencies and other execution requirements are defined by the project setup.

## Project Objective

The primary objective is to transform a repetitive, manual smartphone product entry process into a more structured and automated workflow, supporting efficient product data preparation for Retailix.

## Author

**Rehan Patel**

Computer Science Graduate | Data Analyst Intern

[GitHub Profile](https://github.com/rehaanpatel-29)

---

*Developed to address a practical product management challenge through process automation and structured data handling.*
