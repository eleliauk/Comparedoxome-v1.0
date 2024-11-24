
# Comparedoxome

## Overview
This program is used to compare m/z data from two Excel files, finding matching rows based on setted error tolerance.  Eventually, three output files are generated:
	1. Common.xlsx: contains the data in both normal and cancer data.
	2. Normal only.xlsx: contains the data only in the normal file but not in cancer.
	3. Cancer only.xlsx: contains the data only in the cancer file but not in normal.

## File structure

	-Input file:
	-normal.xlsx: Excel file containing normal data in which the first column is m/z data.
	-cancer.xlsx: Excel file containing cancer data in which the first column is m/z data.

	-Output file:
	-Common.xlsx: contains the data in both normal and cancer data.
	-Normal only.xlsx: contains the data only in the normal file but not in cancer.
	-Cancer only.xlsx: contains the data only in the cancer file but not in normal.



## Parameter description

	-Error tolerance: When the absolute value of the difference between the m/z values in the two files is less than or equal to the setted error tolerance, the lines are considered matched.

## Environmental conditions

```shell
pip install -r requirements.txt
```

## Steps to use

	1. Place normal.xlsx and cancer.xlsx in the data folder.
	2. The default error tolerance is 0.0030 Da.  If necessary, adjust the error tolerance.
	3. Run the program.
	4. After the script runs, it will generate three output files in the data folder: matched.xlsx, maintaining_normal.xlsx, and maintaining_cancer.xlsx.


Common problems

	1. 	No file is generated: Check if there is an input file in the data folder.
	2. 	Result file is empty: The tolerance error may be too small.  Try to increase the value.


