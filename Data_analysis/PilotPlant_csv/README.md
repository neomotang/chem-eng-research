# Analysis of time-series .csv data from a pilot plant facility

---

## Overview
In this project, data from a pilot plant facility was retrieved from time stamped .csv files and summarized in a single spreadsheet. During experiments, a single experimental run started once at least 5 minutes of steady state had been established and ended once another 5 minutes of steady state had passed. The start and end times were recorded manually in a notebook. For each run, .csv files with recordings of temperatures, pressures and flow rates were retrieved from the pilot plant's HMI. This `VBA` code uses the manually entered start and end times to calculate and record the appropriate average properties at steady state. An example of the .xlsm file used (AnalyseCSV.xlsm) is included.
