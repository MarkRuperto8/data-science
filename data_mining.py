### Step 1: Ensure that the following Python packages are imported into your working environment: pandas, numpy, matplotlib, and os.

import os
import pandas as pd
import numpy as np

### Step 2: Setup your working directory.

os.chdir('/Users/markruperto/Desktop/School/GCU/BIT-446/Topic 1')

os.getcwd()

### Step 3: Import the attached "Churn-1" Excel file into a pandas dataframe called "Churn;" add a record field starting at 0.

Churn = pd.read_excel("Churn-1.xlsx")

Churn["record"] = range(len(Churn))

### Step 4: Provide header information regarding the dataframe, including the data types for each column.

print(Churn.head())
print(Churn.info())
print(Churn.describe())

### Step 5: Update all column names in the dataframe to be lower case. Show the "before" and "after" results of this operation.

print("Before:")
print(Churn.columns.tolist())

#Convert str to lowercase
Churn.columns = Churn.columns.str.lower()

print("\nAfter:")
print(Churn.columns.tolist())

### Step 6: Remove special characters and spaces from all column names in the dataframe. Show the "before" and "after" results of this operation.

print("Before:")
print(Churn.head())

#Remove spaces and special characters
Churn.columns = Churn.columns.str.replace(r'[^a-zA-Z0-9]', '', regex=True)

print("\nAfter:")
print(Churn.columns.tolist())

### Step 7: Remove duplicate rows (if any). State how many duplicate rows (if any) where removed.

duplicate_count = Churn.duplicated().sum()

#Removes duplicates
Churn = Churn.drop_duplicates()

print("Number of duplicate rows removed:", duplicate_count)

### Step 8: Count and indicate the missing values that exist in the dataframe.

#Checks every cell in dataframe... True = 1, False = 0
missing_values = Churn.isna().sum()

print("Missing values per column:")
print(missing_values)

#Total missing values
total_missing = missing_values.sum

print("\nTotal missing values:", total_missing)

### Step 9: Detect and impute missing data for continuous variables with the median of the respective column. Then, check to ensure that no missing data exist.

# Step 1: Identify numeric columns
numeric_cols = Churn.select_dtypes(include=['number']).columns
print("Numeric columns:", numeric_cols.tolist())

# Step 2: Impute missing values with median safely
for col in numeric_cols:
    median_value = Churn[col].median()
    Churn[col] = Churn[col].fillna(median_value)

# Step 3: Verify no missing values remain
missing_values_after = Churn.isna().sum()
print("\nMissing values per column after imputation:")
print(missing_values_after)

total_missing_after = missing_values_after.sum()
print("\nTotal missing values after imputation:", total_missing_after)

### Step 10: There is an extreme outlier in the "vmailmessage" column in the dataframe. You are requested to use the "outliers" package to locate the outlier. Complete the following:

from outliers import smirnov_grubbs as grubbs

s = Churn["vmailmessage"].dropna()

#Run Grubbs test (this confirms statistical significance)
grubbs.max_test(s.values)

#Get position of maximum value
pos = np.argmax(s.values)

#Get record number from original dataframe
record_number = s.index[pos]

#Get the actual outlier value
outlier_value = s.iloc[pos]

print("\nOutlier value:", outlier_value)
print("Record number:", record_number)

### Step 11: Recode the "churn" field to remove the "periods." Show the "before" and "after" results of this operation.

print("Before:")
print(Churn.churn.head())

#Regex=False because Regex=True a dot (.) means any character and not just a period
Churn["churn"] = Churn["churn"].str.replace(".", "", regex=False)

print("After:")
print(Churn.churn.head())

### Step 12: Install the "skimpy" package. Using the appropriate command from this package, summarize final Churn dataframe.

from skimpy import skim

#skim() generates a compact, readable summary the entire DataFrame.
skim(Churn)

 ### Step 13: Export final Churn dataframe to an Excel file called "ChurnREADY."

#Add index=false to prevent writing an extra column in excel
Churn.to_excel("ChurnREADY.xlsx", index=False)
