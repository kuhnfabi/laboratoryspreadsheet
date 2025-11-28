Instructions for installation:

INSTALLATION (Windows)
----------------------

1. Install Anaconda ( https://www.anaconda.com/ )
2. Get the LaboratorySpreadSheet  package:
	a. Download from GitLab: LaboratorySpreadSheet 
	b. Unzip into a folder of your liking, e.g. C:/code/laboratoryspreadsheet
3. Open command prompt (e.g. Anaconda Powershell Prompt) and execute the following steps:
	a. Create virtual environment with this command:
		conda create -n LABENV python=3.7 anaconda
	b. Install required packages with this pip command:
		pip install -r C:/G/code/laboratoryspreadsheet/labspreadsheet
4. You are ready!


NORMAL USE
----------

1. [Excel] Ensure you have access to all folders listed in the sheet `file` in `MetaData.xlsx` - e.g.: At Eawag, make sure you are connected to the shared Q drive.
2. [Python] Execute `ScriptA_CopyFromRemote.py`. This will make a copy of all existing Excel sheets and place them in the `.\local_input\` folder. Note that this removes all files in the `.\local_input\` folder first.
3. [Excel] Update the file `MetaData.xlsx` in the folder `.\meta\` to list all project-year combinations for which a laboratory file is in use or should be made. Warning: Do not remove any project unless you remove all projects for the same year at the same time (Doing so will result in loss of data in the central file LabSamplesYYYY.xlsx). 
4. [Excel] Ensure that the sheet `site` in `MetaData.xlsx` lists all feasible sampling locations
5. [Excel] Ensure that the sheet `variable` in `MetaData.xlsx` lists all measured variables that are available
6. [Excel] Check that the sheet `person` and `person_db` in the files with name `LabSamplesYYYY.xlsx` in folder `.\local_input\` is complete. This will be used to define the dropdown menus listing the available staff members.
7. [Python] Execute `ScriptB_CreateNewSpreadSheets.py`. This will make a new copy of all existing Excel sheets and place them in the local `.\output\` folder. The most important elements of this include:
  a. Conditional and regular style formatting is implemented
  b. Data validation is implemented
  c. Implementing links between files
  d. Copying data from existing files
  e. Apply cell locking, hiding specific sheets and columns, and protect workbook and sheets
8. [Excel] Check the produced files in folder `.\local_output\`, manually remove files you do not want to update and adjust the set of visible columns as necessary
9. [Python] Execute `ScriptC_CopyToRemote.py`. This will copy the produced files to the folders indicated in the sheet `file` in `MetaData.xlsx`. This overwrites existing files with the new copies after makeing a copy of the old ones.

CREATE LABSHEETS FOR A NEW YEAR
----------
1. [Excel] Update the "Year" in "MetaData.xlsx" for all projects with the new year.
2. [Excel] Copy "LabSamples_empty.xlsx" to folder `.\local_input\`. Rename the copied excel sheet with the name "LabSamplesYYYY.xlsx", where YYYY is the new year.
3. [Python] Conduct step 7-9 of NORMAL USE.
