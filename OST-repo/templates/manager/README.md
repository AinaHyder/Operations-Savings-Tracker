# 📊 Operations Savings Tracker

> **One source of truth for tracking, approving, and analyzing cost-saving ideas across teams — from forecast to actuals.**

![Status](https://img.shields.io/badge/status-active-brightgreen)
![License](https://img.shields.io/badge/license-MIT-blue)
![Made With](https://img.shields.io/badge/made%20with-Python%20%7C%20Django%20%7C%20SQL-blueviolet)
![UI](https://img.shields.io/badge/UI-Web%20Dashboard-orange)

---

## 📌 Overview
The **Operations Savings Tracker** is a centralized platform that standardizes the process of:
- Submitting cost-saving ideas,
- Approving/rejecting proposals,
- Tracking savings from forecast to actual,
- Maintaining clean master data for accurate reporting.

**Why?**  
Before this tool, savings tracking was fragmented across Kaizen, ZBB, and Jira portals — leading to inconsistent assumptions, incomparable results, non-standardized definitions, and manual errors.  
Now, all stakeholders work in one transparent, automated system.

---

## ✨ Features

- **Role-based access**
  - **User** → Submit and update projects
  - **Manager** → Review, approve/reject, and compare forecast vs. actual savings
  - **Admin** → Maintain master data and guardrails

- **Standardized Input Forms**
  - Categories, products, factories, investment, parameters — all in a single guided form.

- **Forecast vs. Actual Tracking**
  - Update actual units by phase; see real-time variance and trends.

- **Built-in Governance**
  - Currency conversion, ROI calculations, milestone tracking.

- **Export & Analytics**
  - Filter, export data, and analyze savings across the organization.


## Directory Structure: 
```
app.py
requirements.txt
templates/
├── admin/
│   ├── add_edit_account.html
│   ├── admin_dashboard.html
│   ├── admin_projects.html
│   ├── create_project.html
│   ├── currency.html
│   ├── manage_accounts.html
│   ├── project_category.html
│   ├── project_detail.html
│   ├── projects.html
│   ├── unit_costs.html
│   └── upload.html
├── manager/
│   ├── analytics_dashboard.html
│   ├── manager_dashboard.html
│   ├── milestones.html
│   ├── project_detail.html
│   ├── projects.html
│   └── roi_table.html
├── user/
│   ├── actual_timeline_input.html
│   ├── factory_select.html
│   ├── milestone_view.html
│   ├── model_select.html
│   ├── product_select.html
│   ├── project_details.html
│   ├── project_details_input.html
│   ├── project_parameters.html
│   ├── project_timeline_input.html
│   ├── project_type.html
│   ├── roi_table.html
│   └── user_dashboard.html
├── base.html
├── change_password.html
├── documentation.html
├── login.html
└── welcome.html

```


## 🖼 Workflow Diagram

```mermaid
flowchart TD
    A[User Login] --> B[Submit New Project]
    B --> C[Manager Review & Approval]
    C --> D[User Updates Actuals]
    D --> E[Manager Reviews Variance]
    E --> F[Admin Maintains Master Data]