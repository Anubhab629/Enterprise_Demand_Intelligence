# Enterprise Demand Intelligence Platform

An enterprise-grade, dataset-independent demand intelligence platform designed to transform business data into demand insights, forecasts, and actionable business recommendations.

## Overview

The Enterprise Demand Intelligence Platform is designed to support multiple manufacturing and business organizations by combining historical business data, current market conditions, operational information, customer activity, and competitor-related signals.

The platform follows a modular pipeline architecture so that the input dataset can be replaced without redesigning the core processing system.

The main objective is to move from raw business data to:

- Clean and validated data
- Analytical business insights
- Demand forecasts
- Demand-driving factor analysis
- Business recommendations
- Power BI reporting

## Key Capabilities

- Dataset-independent data processing
- Automated data loading and validation
- Data cleaning and preprocessing
- Data quality checks
- Demand and business analysis
- Demand-signal relationship analysis
- Forecast generation
- Recommendation generation
- Optional AI enrichment with deterministic fallback
- Power BI report generation
- Run-specific workspace isolation
- Artifact validation
- Execution status tracking
- Run manifest generation
- Controlled project cleanup
- Environment and source-code protection
- Dataset replacement support

## Architecture

The platform follows this general flow:

Raw Dataset  
↓  
Dataset Validation  
↓  
Data Cleaning  
↓  
Data Preprocessing  
↓  
Business Analysis  
↓  
Demand Forecasting  
↓  
Recommendation Engine  
↓  
Power BI Integration  
↓  
Business Intelligence Report

Each stage is separated so that individual components can be maintained and validated independently.

## Project Structure

```text
Enterprise_Demand_Intelligence/
│
├── data/
│   ├── raw/
│   ├── processed/
│   └── output/
│
├── powerbi/
│   ├── Demand_Intelligence_Report_<RUN_ID>/
│   ├── data/
│   └── documentation/
│
├── src/
│   ├── analysis.py
│   ├── cleaning.py
│   ├── data_loader.py
│   ├── data_profiler.py
│   ├── excel_exporter.py
│   ├── forecasting.py
│   ├── preprocessing.py
│   ├── recommendation_engine.py
│   ├── validation.py
│   └── __init__.py
│
├── config.py
├── env.py
├── main.py
├── powerbi_integration.py
├── project_cleaner.py
├── .gitignore
└── README.md