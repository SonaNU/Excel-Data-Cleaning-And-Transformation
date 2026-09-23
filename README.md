# Excel-Data-Cleaning-And-Transformation

**Project Overview**

This project demonstrates the process of cleaning, transforming, and standardizing a raw product dataset using Microsoft Excel.
The dataset contains product information such as Product ID, Product Name, Brand, Price, Quantity, and Category. The objective of this project is to identify and handle data-quality issues and transform the raw dataset into a cleaner and more structured format suitable for analysis.
This project focuses on practical Excel data-cleaning techniques that are commonly used in Business analysis.

**Objectives**

- Identify and handle missing values
- Correct inconsistent and incorrect data
- Standardize text values
- Remove duplicate records
- Split and transform Product ID information
- Create new derived columns
- Apply appropriate number and date formatting
- Use conditional formatting to improve data visualization
- Prepare the dataset for further analysis

**Data cleaning tasks**

**1. Handling Missing Values**

Missing Price and categories : 
The missing values in the dataset would be marked as Missing through the Go To dialog box, where all the blanks could be filled up.

Blank values in the Price column were identified using Excel's:

**Go To → Special → Blanks**

**2. Correcting Inconsistent Data**

•	Inconsistent product name and Typos in the category column could be corrected through Find and Replace.

•	Also Inconsistent product name could be standardised using the proper()

**3. Removing Duplicate Records**

Duplicate records were identified and removed using:

Data → Remove Duplicates

The duplicate check was performed using all relevant columns.
This helped ensure that the cleaned dataset did not contain repeated records.

**4. Data Transformation**

**Product ID Transformation**

The Product ID contains day,month and country code.

Country Code is extracted by:

=RIGHT(A2,2) and drag the rest of the cells.

Manufacturing Date is extracted by:

=LEFT(A2,6) and drag the rest of the cells.

**Creating a Product Brand Column**

A new derived column called Product Brand was created by combining Brand Name and Product Name.

=CONCAT(C2," ",G2)

**5. Number Formatting**

**Price**

The Price column was formatted as currency using:

Home → Number → Currency

**Manufacturing Date**

Select the manufacturing date cell and format with the short date option to format by the "DD-MM-YYYY " format through the formula

=DATEVALUE(LEFT(A2,6)&"-2026")

**6. Conditional Formatting**

**Price Data Bars**

Data bars were applied to the Price column to visually compare product prices.

Path:

Home → Conditional Formatting → Data Bars

This provides a quick visual representation of relatively low and high prices.

**Highlighting Electronics**

The Electronics category was highlighted using:

Home  Conditional Formatting → Highlight Cells Rules → Text that Contains

**Data Cleaning Workflow**

Raw Dataset → Identify Missing Values → Correct Inconsistent Data → Standardize Text → Remove Duplicates → Split Product ID → Create Derived Columns → Format Data → Apply Conditional Formatting → Clean Dataset

**Conclusion**

This project demonstrates a complete basic data-cleaning and transformation workflow using Microsoft Excel.
The raw product dataset was cleaned, standardized, transformed, and formatted to make it more consistent and suitable for further business analysis.
The project also demonstrates practical Excel skills that can be applied to real-world tasks.
