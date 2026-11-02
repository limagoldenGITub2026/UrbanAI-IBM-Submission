# UrbanAI - Automated Statutory Due-Diligence & Feasibility Platform

UrbanAI is an AI-powered decision support platform that automates statutory data retrieval, building rights analysis, and planning risk assessment for urban renewal projects in Israel.

## 🏗️ Architecture & Technology Stack
- **IBM watsonx Orchestrate**: Central orchestrator coordinating multi-agent workflows.
- **IBM Bob**: Used during development for prompt engineering, agent setup, and task logic flow design.
- **Multi-Agent Architecture**: Dedicated single-purpose agents (Data Collection, Source Verification, Statutory Analysis, Risk Registry, and Reporting).

## 📁 Repository Structure
- `excel_interface/`: Python package handling data reading, mapping, validation, logging, and XPlan statutory API connector.
- `excel-interface-plan.md`: Step-by-step technical development specifications.
- `analysis.md`: Complete field dictionary, class structures, and schema mapping rules.