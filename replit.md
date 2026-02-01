# Overview

This is a Streamlit-based document generation application that processes MP (Member of Parliament) correspondence data. The application takes spreadsheet data containing letters to MPs and generates formatted Word documents organized by electorate. It handles date parsing, MP name substitution, and creates downloadable DOCX files for each electorate's correspondence.

**Current Status**: New application (`new_app.py`) created as a fresh version based on the working system. Both versions are available - the original `app.py` and the new `new_app.py` which runs as "New MP Document Generator".

# User Preferences

Preferred communication style: Simple, everyday language.

# System Architecture

## Frontend Architecture
- **Streamlit Web Framework**: Single-page application using Streamlit for rapid prototyping and simple UI components
- **File Upload Interface**: Handles CSV/Excel file uploads for MP and letter data processing
- **Download System**: Generates and serves ZIP files containing organized DOCX documents

## Backend Architecture
- **Document Processing Pipeline**: 
  - Parses input data from spreadsheets
  - Groups letters by electorate
  - Applies MP-specific formatting and salutations
  - Generates Word documents with proper date sorting
- **Date Handling**: Flexible date parsing supporting multiple formats (MMM DD, YYYY and MMMM DD, YYYY)
- **Template System**: Dynamic placeholder replacement for MP names and salutations in letter content

## Data Processing
- **Pandas Integration**: Used for efficient data manipulation and CSV/Excel file processing
- **Document Generation**: python-docx library for programmatic Word document creation
- **File Organization**: Groups correspondence by electorate and sorts chronologically

## Design Patterns
- **Factory Pattern**: Document creation based on electorate groupings
- **Template Method**: Standardized letter formatting with variable MP information
- **Data Transformation Pipeline**: Multi-stage processing from raw data to formatted documents

# External Dependencies

## Core Libraries
- **Streamlit**: Web application framework for user interface
- **Pandas**: Data manipulation and analysis for spreadsheet processing
- **python-docx**: Microsoft Word document generation and formatting
- **Python Standard Library**: datetime, io, zipfile, re, typing modules

## File Format Support
- **Input Formats**: CSV and Excel files for data import
- **Output Formats**: DOCX (Microsoft Word) and ZIP file generation
- **Data Structure**: Expects structured data with columns for MP information, dates, and letter content

## System Requirements
- **Python Runtime**: Compatible with standard Python 3.x environments
- **Memory Management**: In-memory file processing using io.BytesIO for document generation
- **No Database**: File-based data processing without persistent storage requirements