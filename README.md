# ReportsApplication

An application that allows the user to choose a report, and the application will dynamically add datasets, datasources, parameters, etc.
so that it can render the report locally. It can also output the report as PDF, Word, or Excel. It can combine multiple reports
into a single PDF. May not work with your report if it uses a query that uses a "select not null" statement or if it uses
MDX and the fields don't use standard MDX names.

The meat of the code is in Form1.vb
