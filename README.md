# 💒 Wedding Seating System

Ultra Advanced Production-Ready Wedding Seating System for Google Sheets + Apps Script

## 📋 Overview

This project provides a comprehensive wedding seating management system built on Google Sheets with powerful Apps Script automation. It features real-time seat assignment tracking, conflict detection, VIP priority handling, and intelligent auto-seating algorithms.

## 🚀 Features

### Basic Version (`WeddingSeating_Code.gs`)
- Real-time seat assignment via onEdit trigger
- Master guest list synchronization
- Table header status updates
- Basic validation and error handling

### Ultra Advanced Version (`WeddingSeating_UltraAdvanced_Code.gs`)
- VIP priority seating
- Dietary requirement tracking
- Plus-one management
- Comprehensive logging
- Custom menu integration

### Level 100 Version (`WeddingSeating_LEVEL100_Code.gs`)
- AI-Powered Compatibility Scoring Engine
- Social Network Analysis & Relationship Graph
- Genetic Algorithm for Optimal Seating
- Conflict Detection & Resolution System
- Multi-Objective Optimization
- Real-Time Analytics Dashboard
- QR Code Generation & Check-In
- Advanced Dietary Clustering
- VIP Protocol Engine
- Accessibility Compliance
- Multi-Language Support
- Comprehensive Audit Logging
- Predictive Analytics
- Auto-Recovery System

## 📊 Sheet Structure

| Sheet Name | Purpose |
|------------|---------|
| Guest List 1 | Family guests |
| Guest List 2 | Friends guests |
| Guest List 3 | Colleagues guests |
| Master_Guests | Consolidated guest list with assignments |
| SEATING_PLAN | Visual seating arrangement (20 tables × 10 seats) |
| CONFIG | System configuration |
| DASHBOARD | Analytics and statistics |
| _DEBUG | Validation issues |
| _SCRIPT_LOG | Action history |

## 🔧 Installation

1. Open your Google Sheet
2. Go to **Extensions > Apps Script**
3. Delete any existing code
4. Paste the code from the desired `.gs` file
5. Save (Ctrl+S)
6. Return to sheet and refresh

## 📖 Usage

### Custom Menu
After installation, access via **💒 Wedding Seating** menu:
- ✅ Validate All Data
- 🔄 Refresh All Tables
- 📊 Update Dashboard
- 🔄 Sync Guest Lists
- 🪑 Auto-Seat Unassigned
- 🗑️ Clear All Assignments
- 📋 Generate Report

### Manual Seating
Simply type a guest name in any seat cell (Column B) in the SEATING_PLAN sheet. The system will:
1. Validate the guest exists in Master_Guests
2. Check for conflicts
3. Update the Master_Guests assignment
4. Apply color coding
5. Update table status

## 🎨 Color Coding

| Source | Color |
|--------|-------|
| Family | #1A73E8 (Blue) |
| Friends | #EA4335 (Red) |
| Colleagues | #34A853 (Green) |

## 📈 Google Spreadsheet

[Open Wedding Seating Sheet](https://docs.google.com/spreadsheets/d/1uJrBcIwh3Tc-be8Mt_Xez9xwdDcu-b3ol_zCSIJ5LZM/edit)

## 📁 Files

| File | Description |
|------|-------------|
| `WeddingSeating_Code.gs` | Basic production-ready version |
| `WeddingSeating_UltraAdvanced_Code.gs` | Enhanced version with VIP handling |
| `WeddingSeating_LEVEL100_Code.gs` | Ultimate version with AI optimization |

## 🤝 Contributing

Feel free to fork and enhance this system for your wedding!

## 📜 License

MIT License - Use freely for personal or commercial purposes.
