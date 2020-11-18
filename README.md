# ICPMS_Sample_List
This project makes use of VBA and python macros to create a sample list from a Laboratory Management Information System (LIMS) backlog.  The code is written to import sample information from a Labworks LIMS, but could probably be tailored to other types of LIMS.  The sample information is processed in various ways to create a coversheet, analyte lists, and a suitable sample list to be copied into the instrument software.

## Getting Started
The macro-enabled workbook "ICPMS_Sample_List" (.xlsm) is what the user primarily interacts with.  Within are a series of worksheets used chronologically to process the sample information.  One worksheet, "ElementsBySample", uses the python script (via xlwings) to generate a more user-friendly analyte list.  The workbook is provided for reference, but users will likely have to install xlwings, then use the xlwings quickstart command to create their own project.  See here: https://docs.xlwings.org/en/stable/udfs.html#udfs

Users without access to LIMS can use the provided file "backlog.dat" with test sample information data to get a sense of how the macros work.  The directory of dat file will need to be specified in visual basic.  Alternatively, the excel workbook "ICPMS_Sample_List" has a worksheet "TestData" with data that can simply be copied to the worksheet "Import".

#### Prerequisites

The functionality of these tools assumes that the user's instrument software is compatible with excel clipboard objects (i.e. data can be pasted from excel).

#### Other Considerations

Laboratory analysts usually have very specific requirements for how data is presented in a sample list.  Users will likely have to customize output for thier individual needs.
