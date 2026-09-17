# Project for processing Aglink results, data and the master file for the Medium-term Outlook (MTO)


## Workflow you need to follow with every new run

- copy the merge file (xlsx) from the shared drive to the local folder 'mergefiles'
- rename the merge file to reflect the time of the run (usually day and exact time of the run)
- create a new viewer file by copy/pasting the last version of the viewer and renaming
- adjust filenames in process_Aglink_results.R; the two variables you need are: merge_file and viewer_file
- run the process_Aglink_results.R script
- the script generates a new sheet in the viewer; you still need to manually copy the content of this sheet to the sheet 'BASELINE'
- the .R script also generates an Excel file with the filtered results 'to_copy...' which is intended to archive each steps. This file is not directly used in the update process
- voila, your viewer is updated to the latest results from the merge
- if needed, you can adjust/extend code lists; for this, have a lok at filter_results.R


## Files

-   process_Aglink_results.R: filters the big results cube and prepares a smaller subset to be copied to the viewer
-   milk_yields.R: checks milk yield trends in the Master File
-   complete_balancesheet.R: performs additional calculations to get the missing items in the balance sheet we publish (e.g. waste and FO for the individual Fresh Dairy Products)

