# Common data errors and solutions
# High-level overview
This reading details various common data errors and presents practical solutions using Copilot in Excel. It includes step-by-step instructions on identifying and fixing issues like missing values, inconsistencies, and incorrect data types, enhancing learners' ability to maintain data quality. For instance, a sales manager correcting non-descriptive column headers to ensure clear communication in sales reports. If you would like to try the solutions yourself, you can download the dataset referenced in the additional resources section at the bottom of this reading and use it for practice.

**Learning objective**
By the end of the reading, you will be able to:

identify common data errors, clarify datasets by addressing these errors, and convert data into suitable formats for accurate analysis

**Introduction**
The Marketplace Sales Data dataset includes transaction data from an online store. Unfortunately, the legacy point of sale reporting system for this store misses some entries. In this interactive reading let’s see if we can work on this dataset with the help of Copilot.


**Error 1: Missing values**
The first type of error you should look for in your spreadsheet is missing values. This happens, if for instance, a respondent skips a question in a survey, resulting in the exported data displaying an empty cell where that response should have been. It’s quite a common occurrence but, to make the most of your data, you will need to address this issue effectively. Otherwise, missing values can disrupt calculations, complicate predictive modeling, hinder data merging or comparison, skew findings, and lead to incomplete or biased conclusions.

**Solution using Copilot**
Detecting missing values

In Copilot, type Highlight cells with missing data.

Wait a few seconds for Copilot to generate a response, then select Apply.

You will see that the empty cells are now highlighted.

**Handling missing values**

Type Suggest actions for missing data into Copilot.

Review the list of suggested actions and choose the method that best fits your data and analysis needs. I recommend starting with Identify the missing data.

Ask Copilot What header do the missing values come under? It will respond: Total revenue.

Ask Copilot What is the formula for total revenue? It will reply: Unit price multiplied by units sold.

To calculate total revenue, tell Copilot to Calculate total revenue. Copilot can create a new column for this calculation; simply select Insert Column. Alternatively, you can manually enter the formula. In cell G2, type =E2*F2 and drag the formula down the column.

After completing these steps, review the data to ensure accuracy and consistency.

**Error 2: Inconsistencies**
The second type of error you can very easily run into is inconsistencies. These errors manifest in various forms, such as duplicate entries, different naming conventions like using "BC" instead of "British Columbia" or mixing measurement units such as kilograms versus pounds. Again, this issue is particularly common in surveys and requires some fiddly checks to ensure you get the most out of your analysis. If left unresolved, inconsistencies can make aggregating and summarizing information quite troublesome, resulting in misleading insights and inaccurate reports. Inconsistencies can also hinder comparisons across datasets, cause unexpected issues in data processing, and ultimately undermine the trustworthiness of your data.

**Solution using Copilot**
Identifying data inconsistencies

In Copilot, type Check for inconsistencies in the data.

Review the list of available inconsistency checks.

Select an option, such as Duplicate Entries.

Type Highlight rows with duplicated data. Select Apply.

The rows with duplicate data will be highlighted.

**Handling duplicate entries**

In Copilot type Remove duplicates. 

Copilot will guide you through the steps to manually remove duplicates, as it cannot perform this action directly.

Follow these instructions: Go to the Data tab, select Data Tools, and click Remove Duplicates.

Choose the column(s) where you want to remove duplicates. For example, select Transaction ID because it is a unique identifier, unlike Product Name, where multiple entries might exist for the same product.

Click OK to remove the duplicates.

Alternatively, you can tell Copilot to Delete rows 3 and 6 if you know the specific rows to remove.

Review the data to ensure that duplicates have been successfully removed.

**Error 3: Incorrect data types**
The third type of error you're likely to encounter is incorrect data types. For example, in a survey, respondents may write "n/a" (not applicable) for questions that don’t apply to them. When non-numeric entries like “n/a" appear in cells meant for numeric values, such as hours of exercise or number of cars owned, this can be a bit of a problem. You could end up facing difficulties with calculations or having distorted data, skewed results, and poorly represented graphs.

**Solution using Copilot**
Detecting incorrect data types

Ask Copilot Do I have any incorrect data types?

Copilot will explain how to check for invalid values. To do this:

Select the column in question, such as column E (Units Sold).

Go to the Data tab on the Ribbon and select Data Validation. 

In the Data Validation dialog box, under the Settings tab, choose Whole Number from the Allow list.

Set a minimum and maximum range, such as 1 and 10, and press OK.

Cells with incorrect data types will have little corner notes indicating the issue. For example, cell E4 contains "one" written as text instead of a whole number, so it is flagged. 

To highlight these cells in color, type Show me cells with invalid data types in column E into Copilot.

**Handling incorrect data types**

Manually change text-based numbers to numeric values in the highlighted cells, for example, replace "one" with "1", "two" with "2", and so on.

**Conclusion**
Like any new tool, Copilot takes some getting used to, but its ability to automate tasks quickly can make managing large datasets a whole lot easier. 

Practice with Copilot and you’ll find that the efficiency it brings to your data management is well worth the effort.

Feel free to try out these techniques with your datasets and put your newfound knowledge into action. 
