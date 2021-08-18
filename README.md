# Data Aggregation

The main idea of this script is to aggregate, slice and store new inventory reports coming everyday.

The general input file in this script is a .scv, but it can be any dataframe with stock and sales data per product and DC. Later data is sliced into different Excel sheets according to DC names. One sheet is kept for the original dataframe with all DCs.

![Sheets in the output file.](https://www.screenpresso.com/=lGVoc)

As the result, we get an Excel file with processed inventory reports and a log file which stores the processing duration for every run and total number of lines in the general dataframe. When the duration time becomes significantly longer, in the next iteration I'm going to improve the script.

![File with logs.](https://www.screenpresso.com/=czhkc)


By now the script is launched by Task Scheduler everyday at 4 PM.

I would mark this whole process as the first step towards the topic of time series forecasting, which I'm interested in.
