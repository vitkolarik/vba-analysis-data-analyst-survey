# vba-practice-project
This is my final semestral assigment for subject "Economic modeling in a spreadsheet with VBA" I took in autumn 2022. 

## Goal
The goal of this work was to upload the source data from the CSV file into MS Excel and then do various changes like clean and prepare the data, delete unnecessary columns, make summary tables from selected questions and than create graphs with visualizing the results. All of these changes are written in single VBA macro and are done at the same time. 
This macro could be used e.g. for regular monthly analysis with standartized form of source data. 

## Data
The source data I've used are from the "Power BI tutorial" on YouTube channel "Alex The Analyst". Can be found at these links.
https://youtu.be/pixlHHe_lNQ
https://github.com/AlexTheAnalyst/Power-BI/blob/main/Power%20BI%20-%20Final%20Project.xlsx

## Viewing the code
You can either view just the plain VBA code in text form here in GitHub, or download the Excel file with macro and run in in your computer. For that you will need to download the "vba-practice-project.xlsm" with the macro itself and "vba-practice-project-data.csv" containing the source data. These files have to be in the same folder for the code to run succesfuly. 

After you download the "vba-practice-project.xlsm" file, you might need to right/click on the file, go to "Properties" and in "Security" click "Unblock". This might be blocking the macro from running. For running the macro you need the developers tab in your Excel. In case you don't have it, go to Files - Options - Customize Ribbon - Main Tabs - select the Developer check box.

## Running the macro
Once in "vba-practice-project.xlsm", go to the Developers tab and Macros. You will see to two option - "execute" and "reset_to_original". The "execute" macro will run the code and do all the changes. If you would like ot run the code again, the "reset_to_original" macro will delete all changes and you can run the code again. 
