#Radar Chart Plotting Script

This script generates a radar chart from an Excel file and supports interactive annotations and saving the chart as an image.


**Prerequisites**

Ensure you have the following libraries installed:

- pandas
- numpy
- matplotlib
- mplcursors

You can install these libraries using pip: pip install pandas numpy matplotlib mplcursors


**Excel File Format**

The Excel file DataToPlot.xlsx provides the expected structure of the spreadsheet, with the properties section in the first 10 rows and the data section starting from row 11


**Usage**

Ensure your Excel file DataToPlot.xlsx is properly formatted as described above.
Run the script: python radargraphmaker.py


The script will:

Read the data and properties from the Excel file.
Generate a radar chart based on the data.
Annotate data points as specified in the Excel file.
Provide an interactive chart with hover functionality.
Save the chart as an image if specified in the Excel file.


**Features**

Interactive Annotations: Hover over data points to see annotations.
Customizable Appearance: Define colors, line types, markers, and more in the Excel file.
Save Chart: Optionally save the generated chart as an image.


**Notes**

Ensure the Excel file path is correctly specified in the script (file_path variable).
Customize the chart appearance by adjusting the properties in the Excel file.


For further customization, refer to the matplotlib and mplcursors documentation.
