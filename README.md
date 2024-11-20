

# Comparedoxome

## **Overview**

This script compares the m/z data from two Excel files (normal.xlsx and cancer.xlsx) based on a specified error tolerance (tolerance). The output includes three Excel files:

​	1.	matched.xlsx: Contains matched rows from both normal and cancer data.

​	2.	remaining_normal.xlsx: Contains rows in normal data that have no matches.

​	3.	remaining_cancer.xlsx: Contains rows in cancer data that have no matches.

## **File Structure**

**Input Files**

​	•	normal.xlsx: Contains the normal data with m/z values in the first column.

​	•	cancer.xlsx: Contains the cancer data with m/z values in the first column.

**Output Files**

​	•	matched.xlsx: Matched rows from normal and cancer files, sorted by m/z.

​	•	remaining_normal.xlsx: Rows in normal that do not have matches.

​	•	remaining_cancer.xlsx: Rows in cancer that do not have matches.



**Parameters**



​	•	tolerance: The error range for matching m/z values. If the absolute difference between two m/z values is less than or equal to this value, they are considered a match.



## **Dependencies**



Ensure you have the required Python libraries installed:


```shell
pip install pandas openpyxl
```


## **Usage Steps**



​	1.	Place the input files (normal.xlsx and cancer.xlsx) in the data folder.

​	2.	Adjust the tolerance value in the script if needed.

​	3.	Run the script:


```shell
python main.py
```


​	4.	After execution, the output files will be generated in the data folder:

​	•	matched.xlsx

​	•	remaining_normal.xlsx

​	•	remaining_cancer.xlsx



**Output File Structure**



**matched.xlsx**



​	•	Contains matched rows from normal and cancer.

​	•	Each row combines data from both files. Column names are prefixed with normal_ and cancer_ for clarity.



**remaining_normal.xlsx**



​	•	Lists rows from normal.xlsx that do not match any rows in cancer.xlsx.



**remaining_cancer.xlsx**



​	•	Lists rows from cancer.xlsx that do not match any rows in normal.xlsx.



**Notes**



​	1.	**Error Tolerance**: The matching process depends on the tolerance parameter. A smaller value results in fewer matches, while a larger value increases the match count. Adjust as needed.

​	2.	**Empty Sheets**: If no matches or no unmatched rows exist, the corresponding Excel file will contain an empty sheet.



**Example Output**



After running the script, you will see the following files in the data folder:

​	•	matched.xlsx: Matched rows with combined data from both files.

​	•	remaining_normal.xlsx: Rows in normal.xlsx without matches.

​	•	remaining_cancer.xlsx: Rows in cancer.xlsx without matches.



**Script Workflow**



​	1.	**Data Loading**: Reads the normal.xlsx and cancer.xlsx files and extracts the m/z values.

​	2.	**Matching Process**: Iterates through m/z values to find matches within the tolerance range.

​	3.	**Output Generation**: Saves matched rows and unmatched rows to their respective Excel files.



**Common Issues**



​	1.	**No Output Files**: Ensure the data folder contains the input files (normal.xlsx and cancer.xlsx).

​	2.	**Empty Match Results**: Try increasing the tolerance value if no matches are found.



This script simplifies the comparison of research data, particularly when working with values requiring error-tolerance-based matching.