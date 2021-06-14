# Excel_Splitter
Quick program to split excel files by sheet. Provide the program with the directory that contains your excel files that you want to split, and it will split all workbooks in your folder, and create a new folder titled "Split_Results" (in the directory that you provided) that houses the new excel files. New files will use the naming convention oldFileName_sheetName.

Uses tkinter for a gui. Main logic is a for loop within a for loop. Outer loop iterates over all excel files in a folder, inner loop iterates over every sheet in each excel file. 
