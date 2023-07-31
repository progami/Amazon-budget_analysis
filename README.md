# Amazon Advertising Budget Allocator

This program analyzes the performance of your Amazon advertising campaigns and calculates an optimal budget allocation based on past performance. It then updates a bulk file that can be uploaded to Amazon to implement these changes.

## Data Sources

### Ad Group Statistics (AdGroupStats.csv)

This file contains detailed performance metrics for each ad group over a certain period. 

You can download this data from **Scale Insights > Ads Insights > Performance > Ad Groups**. Select an appropriate date range (for example, the last 30 days) and download the CSV file. Save this file as `AdGroupStats.csv` in the working directory of this script.

### Bulk File (bulk_file.xlsx)

This is a template that can be used to make changes to multiple campaigns at once.

You can download this file from **Amazon Campaign Manager > Bulk Operations**. Set 'Yesterday' as the date range, and check 'campaign items with zero impressions'. Save this file as `bulk_file.xlsx` in the working directory of this script.

## Configuration

You can configure the script by changing the values in the `configs` dictionary:

- `MIN_SPEND_NORMAL_CAMPAIGNS`: Minimum daily spend for normal campaigns.
- `MIN_SPEND_TESTER_CAMPAIGNS`: Minimum daily spend for tester campaigns.
- `MIN_SPEND_SP_MANUAL_RANKING`: Minimum daily spend for SP Manual Ranking campaigns.
- `CLICK_COUNT_THRESHOLD`: Minimum number of clicks to be considered a non-tester campaign.
- `NUM_DAYS`: The number of days over which to spread the total ad spend.
- `MAX_SPEND_SD`: Maximum daily spend for SD campaigns.
- `MAX_SPEND_SB`: Maximum daily spend for SB campaigns.

## Running the Script

You can run this script using Python 3 as follows:

```sh
python3 budget_analysis.py
```

## Output

The script generates an updated bulk file (`updated_bulk_file.xlsx`) that can be uploaded to Amazon to apply the new budget allocations. It also prints the summary of budget allocation to the console.

The script also creates an Excel file (`UpdatedAdGroupStats.xlsx`) with the updated ad group data, including the new daily spend values.

## Note

Always ensure to back up your original files before running this script. This script does not modify the original files, but it's always a good practice to keep a backup.
