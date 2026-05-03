# MIS Attendance Generator

MIS Attendance Generator is a simple utility made to reduce manual work while preparing monthly attendance MIS reports.

It takes the Multiple Check-In Excel file exported from T-Hawk HR, automatically prepares the MIS in the required format, fills missing dates for the full month, applies common attendance logic like holidays, leave, credit days, late check-ins, forgot checkout, and half-days, and allows final review before generating the MIS Excel file.

---

## How to Use

### Step 1: Download Attendance File

Open the T-Hawk HR portal:

https://thawksolution.com/

Go to:

**HR Module → Attendance Reports → Multiple Check-In**

Select the required date range for the month and export the report in Excel format.

This exported Excel file will be used as the source file for MIS generation.

---

### Step 2: Upload File

Open the MIS Utility Tool and upload the exported Multiple Check-In Excel file.

The tool will automatically read attendance data and generate the monthly MIS preview.

---

### Step 3: Review and Edit

Before final generation, you can:

* Select global client name
* Add custom client name
* Change client name date-wise from preview
* Select laptop used (Personal / SKA Laptop)
* Mark leave dates
* Mark credit days
* Remove all auto-generated remarks if needed
* Edit remarks manually for specific dates

This helps ensure the final MIS is accurate and ready for submission.

---

### Step 4: Generate Final MIS

Click **Generate** to create the final MIS Excel file in the required reporting format.

The final file can then be reviewed, downloaded, and shared.

---

## Purpose

The main purpose of this tool is to save time, avoid repetitive manual formatting work, and keep MIS reporting accurate, consistent, and easy to review.
