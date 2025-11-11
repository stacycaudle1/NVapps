# NVapps - Business Application Risk Tracker

Note: This repository contains two branches.
Please use the code branch for all updates — changes will be merged into main after review.

A comprehensive desktop application for tracking and managing business applications, their system integrations, risk assessments, and organizational relationships.

## 🎯 Overview

The NVApps application is a standalone, collaborative platform designed to help business units track, evaluate, and optimize key operational processes and applications. 
By providing a centralized space for monitoring workflows that influence business models, budgets, and strategic decisions (Risk), 
NVApps enables teams to make data-driven choices that enhance efficiency and financial planning. This tool supports better forecasting, resource allocation, 
and performance visibility—empowering leadership to identify trends, anticipate challenges, and align investments with long-term organizational goals.

NVapps is a Python-based GUI application built with Tkinter that helps organizations manage their application portfolio with a focus on:

- **Business Unit Management**: Track applications across multiple business units and divisions
- **Risk Assessment**: Score and monitor risk levels for applications and integrations
- **System Integration Tracking**: Document integration points and dependencies between systems
- **Category Classification**: Organize applications by functional categories
- **Reporting & Analytics**: Generate risk reports, business unit summaries, and export to Excel

## ✨ Key Features

### 📊 Application Management
- Create and manage application records linked to business units, divisions, and categories
- Track multiple business units per application
- Assign multiple categories to each application for flexible classification
- Separate display rows for each business unit-division-category combination

### 🔗 System Integration Tracking
- Document system integrations for each application
- Track integration vendors, criticality scores, and risk assessments
- Automated risk score calculation based on configurable factors
- Integration-level notes and status tracking

### 🎨 Risk Assessment & Visualization
- **Risk Ranges**: Low (1-39), Medium (40-69), High (70+)
- Color-coded visualization (green/yellow/red) throughout the interface
- Business Risk calculation combining multiple risk factors
- Disaster Recovery priority banding

### 📈 Comprehensive Reporting
Multiple report types with filtering and export capabilities:

- **Risk Range Report**: Filter integrations by risk level with grouping by business unit and division
- **Business Unit Risk Overview**: Aggregate risk metrics per business unit
- **Division Risk Overview**: Risk analysis grouped by division/application
- **Category Risk Overview**: Risk assessment organized by functional category

### 📤 Export Capabilities
- Single-sheet XLSX export for individual reports
- Multi-sheet export combining all reports in one workbook
- Formatted Excel output with color-coding and professional styling
- Timestamp and metadata included in exports

### 🎨 Modern UI/UX
- Clean, professional interface with Microsoft-inspired design
- Color-coded risk indicators
- Responsive treeview tables with sorting capabilities
- Modal dialogs for data entry and management
- Real-time validation and duplicate detection
- Animated visual feedback for user actions

## 🚀 Getting Started

### Prerequisites
- Python 3.8 or higher
- Windows, macOS, or Linux

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/stacycaudle1/NVapps.git
   cd NVapps
   ```

2. **Create and activate a virtual environment**
   ```bash
   # Windows
   python -m venv .venv
   .venv\Scripts\activate

   # macOS/Linux
   python3 -m venv .venv
   source .venv/bin/activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

### Running the Application

**Option 1: Direct execution**
```bash
python main.py
```

**Option 2: Using VS Code tasks** (if configured)
- Open the project in VS Code
- Run task: "Run NVapps (venv)"

## 📦 Dependencies

- **reportlab** (4.0.0) - PDF generation capabilities
- **matplotlib** (3.8.0) - Charting and visualization
- **pandas** (2.2.2) - Data manipulation and analysis
- **openpyxl** (3.1.2) - Excel file export functionality

## 🗄️ Database Schema

The application uses SQLite with the following key tables:

- **business_units**: Organizational units
- **divisions**: Application divisions linked to business units
- **applications**: Core application records
- **categories**: Functional category definitions
- **system_integrations**: Integration records with risk scoring
- **application_business_units**: Many-to-many relationship between apps and business units
- **application_categories**: Many-to-many relationship between apps and categories
- **division_categories**: Links divisions to functional categories
- **integration_categories**: Categorizes system integrations

## 📋 Usage Examples

### Adding a New Application
1. Select a Business Unit from the left panel
2. Choose or add a Division
3. Select a Category
4. Click "Submit" to create the application record
5. New entries appear as separate rows without modifying existing data

### Managing System Integrations
1. Select an application from the main table
2. Click "Manage Integrations" button
3. Add integrations with vendor, criticality, and risk details
4. Risk scores are automatically calculated

### Generating Reports
1. Click "Show Reports" in the main window
2. Navigate between tabs: Risk Range, Business Unit Risk, Division Risk, Category Risk
3. Use filters to refine the data (e.g., risk range selection)
4. Click "Export XLSX" for single reports or "Export All (Multi-Sheet)" for comprehensive export
5. Refresh buttons update data with timestamp tracking

## 🔧 Configuration

### Risk Scoring
Risk scores are calculated using the formula:
```
Risk Score = (Criticality × Score) ÷ 100
```

### Risk Thresholds
- **Low Risk**: 1-39 (Green)
- **Medium Risk**: 40-69 (Yellow)
- **High Risk**: 70+ (Red)

### Disaster Recovery Priority
Bands: A, B, C, D, E, F based on business risk assessment

## 🏗️ Project Structure

```
NVapps/
├── main.py                 # Application entry point
├── gui.py                  # Main GUI implementation (6000+ lines)
├── database.py             # Database operations and schema
├── integration_handler.py  # Integration management logic
├── requirements.txt        # Python dependencies
├── business_apps.db        # SQLite database (created on first run)
├── scripts/               # Utility scripts
├── templates/             # CSV import templates
└── .venv/                 # Virtual environment (created during setup)
```

## 🛠️ Development

### Database Utilities
- `check_schema.py` - Verify database schema
- `show_database.py` - Display database contents
- `verify_env.py` - Validate Python environment
- `seed_demo_data.py` - Populate database with sample data

### Migration Scripts
- `category_notes_migration.py` - Migrate category data
- `migrate_applications_risk_disaster.py` - Update risk/disaster fields
- `fix_schema.py` - Schema repair utilities

## 🤝 Contributing

Contributions are welcome! Please follow these guidelines:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is private and proprietary. All rights reserved.

## 👤 Author

**Stacy Caudle**
- GitHub: [@stacycaudle1](https://github.com/stacycaudle1)

## 🐛 Known Issues & Roadmap

### Current Limitations
- Chart functionality temporarily disabled in reports
- CSV import functionality not yet implemented
- No authentication/multi-user support

### Planned Features
- Enhanced charting and visualization
- CSV import for bulk data loading
- Advanced filtering and search capabilities
- Export to additional formats (PDF reports)
- User preferences and saved views
- Audit logging and change history

## 📞 Support

For questions, issues, or feature requests, please open an issue on GitHub.

---

**Note**: This application is designed for internal business use and risk management tracking. Ensure proper data handling and security practices when dealing with sensitive business information.
