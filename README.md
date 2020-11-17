# ICPMS_Sample_List
This project makes use of VBA and python macros to create a sample list from a Laboratory Management Information System (LIMS) backlog.  The code is written to import sample information from a Labworks LIMS, but could probably be tailored to other types of LIMS.  The sample information is processed in various ways to create a coversheet, analyte lists, and a suitable sample list to be copied into the instrument software.

## Getting Started
Users without access to LIMS can use the provided file "backlog.dat" with test sample information data. Alternatively, the excel workbook "ICPMS_Sample_List" has a worksheet "TestData" with data that can simply be copied to the worksheet "Import" in the same workbook.

#### Prerequisites

The functionality of these tools assumes that the user's instrument software is compatible with excel clipboard objects (i.e. excel data can be pasted).

