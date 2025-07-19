# insightboard
Meet Insight Board — your AI-powered spreadsheet sidekick! It pulls data from messy Excel files, applies secret spreadsheet ninja rules 🥷📊, and pastes it like a pro into the right place.  Built with Python, sprinkled with AI logic, and fueled by too much coffee ☕ + .xlsx files.

# Driver Matching & MVR Classification Streamlit App

## Overview

This is a multi-functional Streamlit web application designed to assist with driver data management and Motor Vehicle Record (MVR) violation classification. It features advanced fuzzy name matching techniques to compare driver names from multiple Excel datasets, enrich driver information, and help classify violation descriptions based on a reference violation dataset.

The app is built with user experience in mind, featuring a dark-themed UI and a sidebar navigation menu to switch between various tools.

---

## Features

### 1. All Trans MVR - Driver Matching Tool

- Upload two Excel files:
  - **Driver List**: Contains comprehensive driver data including names, hire dates, DOB, license state, etc.
  - **Output File**: Contains a list of drivers to be matched and enriched.
- Intelligent name matching using advanced normalization and fuzzy matching strategies:
  - Handles initials, prefixes, suffixes, and multiple name formats.
  - Uses fuzzy string matching with token-based and partial ratios for accurate matching.
- Transfers important columns such as Hire Date, Date of Birth (DOB), and License State from Driver List to Output File.
- Marks matched records with `"MATCH FOUND"` in the notes column.
- Appends unmatched drivers from the Driver List at the end of the output file with a note `"MISSING MVR"`.
- Allows users to configure row skipping, sheet selection, and performs fuzzy auto-detection of relevant columns.
- Generates a timestamped downloadable Excel output preserving all sheets from the original output file.

### 2. MVR GPT - Violation Classification

- Load a reference Excel dataset (`violations.xlsx`) containing violation descriptions and their categories.
- Users can input a free-text violation description.
- The app uses fuzzy matching to find the closest known violation description and returns its category (e.g., Minor, Major, Accident).
- Provides confidence scoring and suggestions to consult the QC team for uncertain matches.
- Helps automate the classification process for violation records.

### 3. Placeholder Tools (Coming Soon)

- HDVI MVR Tool
- Truckings IFTA Tool
- Riscom MVR Tool

---

## How to Run the Application

### Prerequisites

- Python 3.9 or higher installed.
- Recommended: Use a virtual environment for dependency management.

### Installation Steps

1. Clone this repository or download the source code.

2. Create and activate a virtual environment (optional but recommended):

   ```bash
   python -m venv venv
   source venv/bin/activate   # On Windows use: venv\Scripts\activate

   # PDF Text Searcher 🔍

A powerful PDF search tool with semantic understanding and precision highlighting.

## Features

- Upload and process PDF documents
- Semantic search with keyword proximity scoring
- Context extraction with highlighted results
- Visual page preview with highlighted matches
- Search history tracking

## Installation

1. Clone the repository:
```bash
git clone https://github.com/your-username/pdf-text-searcher.git
cd pdf-text-searcher
