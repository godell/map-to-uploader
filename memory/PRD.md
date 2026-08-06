# Product Requirements Document (PRD) - SAP WM Transfer Order Processor

## 1. Problem Statement
Build a landing page to process attached SAP Sourcedb data with an initial data upload step. The system automatically generates a new column named **Row** by extracting 2 middle characters from the **Source Storage Bin** column starting from the 4th character. The landing page features interactive range cards displaying unique Transfer Order Numbers filtered by the generated Row column (e.g., Card 1 for Row range 1-5, Card 2 for 6-10). Clicking a card displays the filtered Transfer Order items in a table below, optimized for direct tab-separated copy-pasting into SAP GUI.

## 2. User Personas
- **SAP Warehouse Clerk / Inventory Controller**: Needs quick filtering of transfer orders by storage row and seamless copy-pasting into SAP GUI.
- **Logistics Supervisor**: Uploads daily Sourcedb reports to review bin distribution and transfer order item details.

## 3. Core Requirements & Implementation Status
- **File Upload / Data Source**: Supports uploading .csv / .xlsx files or using built-in sample data (`Sourcedb.xlsx`). *(Implemented)*
- **Automated Row Generator**: Adds a `Row` column extracting 2 characters starting from index 3 (4th char) of `Source Storage Bin`. *(Implemented)*
- **Interactive Card Menu**: Dynamically groups unique Row values into range cards (e.g., batch size 3, 5, 10, 15). *(Implemented)*
- **Table Viewer & SAP Copy-Paste**: Displays filtered Transfer Order data with a one-click copy button formatted in TSV (Tab-Separated Values) for SAP GUI (Ctrl+V). *(Implemented)*

## 4. Mocked in Frontend
- File parsing and data state management are fully operational on local mock state.
- Clipboard copy utilizes browser Clipboard API with TSV formatting.

## 5. Prioritized Backlog & Next Steps (Phase 2)
- **P0**: Connect backend FastAPI REST API with MongoDB for persistent data storage of uploaded Sourcedb datasets.
- **P1**: Add user authentication and role-based access for warehouse zones.
- **P2**: Export processed data directly to Excel/PDF reports.

## 6. Next Action Items
1. User review and acceptance of UI/UX and core processing logic.
2. Phase 2 backend wiring (FastAPI + MongoDB).
