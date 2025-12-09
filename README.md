
# S3D Plant Filter XML Generator

A PyQt6 desktop tool that converts Excel inputs into **Plant Filter XML** files for Hexagon Smart 3D (S3D).  
It supports both **Detail** and **Bulk** modes and outputs `FiltersAndStyleRules`-compliant XML ready for S3D import.

## Features
- **Detail mode**: reads `S3DFilter` sheet with `Name(Filter)`, `FullPath(Filter)`, `ObjectPath`
- **Bulk mode**: combines
  - `1.S3DFilterPath` → `FilterPath(Template)`
  - `2.S3DFilterName` → `Name(Filter)`, `SystemName(WBS)`
  - `3.FixedObjectPath` → `ObjectPath(Template)`
  to generate multiple filters at once
- XML structure includes `Information(FileType=FiltersAndStyleRules)` and `PlantFilters`
- Status updates and error handling dialogs
![Screenshot of the S3D Filter XML Generator application](imagesUI.png)

## UI Flow
- Tabs: **1. Filter (Detail)** / **2. Filter (Bulk)**
- Click **Load** to choose Excel → **Run** to generate XML → Save dialog

## How to Use
1. Launch app, select mode
2. Prepare Excel sheets referencing the provided templates
3. Click **Load** to import Excel
4. Click **Run** → generate XML → choose output path
5. In S3D, import via `Automation Tool Kit > Import Styles, Style Rules and Filters`

## Notes
- Missing required sheets/columns show a clear error dialog.
- In Bulk mode, `SystemName(WBS)` is used as path suffix when available; otherwise falls back to `Name(Filter)`.
- Default `MFObjectType` = `Systems\\PipelineSystems`; `MFSystem` uses `IncludeNested=True`.
- If ISO drawing creation throws “Unexpected error,” check whether the filter actually captures any **PipeRun**

