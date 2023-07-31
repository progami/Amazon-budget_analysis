import pandas as pd
from datetime import datetime
from tabulate import tabulate
from openpyxl import load_workbook
import warnings
warnings.filterwarnings("ignore")


configs = {
    "MIN_SPEND_NORMAL_CAMPAIGNS": 1.5,
    "MIN_SPEND_TESTER_CAMPAIGNS": 3.0,
    "MIN_SPEND_SP_MANUAL_RANKING": 3.5,
    "CLICK_COUNT_THRESHOLD": 10,
    "NUM_DAYS": 30,
    "MAX_SPEND_SD": 3.5,
    "MAX_SPEND_SB": 3.5,
}


def find_column(sheet, target_header):
    for cell in sheet[1]:
        if cell.value == target_header:
            return cell.column

def update_bulk_file(df):
    workbook = load_workbook('bulk_file.xlsx')
    budget_mapping = df.set_index('Campaign')['daily_spend'].to_dict()
    sheets_and_budget_headers = {
        'Sponsored Products Campaigns': 'Daily Budget',
        'Sponsored Brands Campaigns': 'Budget',
        'Sponsored Display Campaigns': 'Budget'
    }

    for sheet_name, budget_header in sheets_and_budget_headers.items():
        sheet = workbook[sheet_name]
        entity_col = find_column(sheet, 'Entity')
        operation_col = find_column(sheet, 'Operation')
        campaign_name_col = find_column(sheet, 'Campaign Name')
        budget_col = find_column(sheet, budget_header)

        for row in sheet.iter_rows(min_row=2):
            if row[entity_col-1].value == 'Campaign':
                row[operation_col-1].value = 'Update'
                campaign_name = row[campaign_name_col-1].value
                if campaign_name in budget_mapping:
                    row[budget_col-1].value = budget_mapping[campaign_name]

    workbook.save('updated_bulk_file.xlsx')


def load_and_preprocess_data(filename):
    df = pd.read_csv(filename)
    df.fillna(0, inplace=True)
    df = df.drop(columns=['Units', 'ROAS', 'Default Bid'])
    return df

def calculate_metrics(df):
    total_spend = df['Spent'].sum()
    total_sales = df['Sales'].sum()

    df['%ad spend'] = df['Spent'] / total_spend
    df['%ad sales'] = df['Sales'] / total_sales

    df['distributed_spend'] = total_spend * df['%ad sales']
    df['daily_spend'] = df['distributed_spend'] / configs["NUM_DAYS"]

    return df

def apply_constraints(df):
    df = df[df['State'] == 'Enabled'].copy()
    df['%ad spend'] = df['%ad spend'].round(4)
    df['%ad sales'] = df['%ad sales'].round(4)
    df['distributed_spend'] = df['distributed_spend'].round(2)
    df['daily_spend'] = df['daily_spend'].round(2)

    df['daily_spend'] = df['daily_spend'].apply(lambda x: max(x, configs["MIN_SPEND_NORMAL_CAMPAIGNS"]))
    df.loc[(df['Type'] == 'SP Manual') & (df['AdGroup'] == 'Ranking'), 'daily_spend'] = df.loc[(df['Type'] == 'SP Manual') & (df['AdGroup'] == 'Ranking'), 'daily_spend'].apply(lambda x: max(x, configs["MIN_SPEND_SP_MANUAL_RANKING"]))

    df.loc[df['Type'].str.startswith('SD'), 'daily_spend'] = df.loc[df['Type'].str.startswith('SD'), 'daily_spend'].apply(lambda x: min(x, configs["MAX_SPEND_SD"]))
    df.loc[df['Type'].str.startswith('SB'), 'daily_spend'] = df.loc[df['Type'].str.startswith('SB'), 'daily_spend'].apply(lambda x: min(x, configs["MAX_SPEND_SB"]))

    df.loc[(df['Clicks'] < configs["CLICK_COUNT_THRESHOLD"]) & (df['Type'].str.contains('SP')), 'daily_spend'] = df['daily_spend'].apply(lambda x: max(x, configs["MIN_SPEND_TESTER_CAMPAIGNS"]))

    return df

def calculate_and_print_results(df):
    daily_tester_budget = df.loc[(df['Clicks'] < configs["CLICK_COUNT_THRESHOLD"]) & (df['Type'].str.contains('SP')), 'daily_spend'].sum()
    daily_other_budget = df.loc[~((df['Clicks'] < configs["CLICK_COUNT_THRESHOLD"]) & (df['Type'].str.contains('SP'))), 'daily_spend'].sum()
    total_daily_budget = df['daily_spend'].sum()

    total_spend = df['Spent'].sum()
    total_initial_spend_daily = round(total_spend/configs['NUM_DAYS'], 2)
    percent_daily_tester_budget = (daily_tester_budget/total_initial_spend_daily)*100
    percent_daily_other_budget = (daily_other_budget/total_initial_spend_daily)*100
    percent_total_daily_budget = (total_daily_budget/total_initial_spend_daily)*100

    results = {
        "Total initial spend (daily)": f"£{total_initial_spend_daily}",
        "Daily budget for tester campaigns": f"£{daily_tester_budget} ({percent_daily_tester_budget:.2f}%)",
        "Daily budget for other campaigns": f"£{daily_other_budget} ({percent_daily_other_budget:.2f}%)",
        "Total daily budget": f"£{total_daily_budget} ({percent_total_daily_budget:.2f}%)"
    }

    print("Results:")
    print(tabulate(results.items(), headers=["Parameter", "Value"], tablefmt="pretty"))

def calculate_spend_percentage(df):
    total_spend_daily = round(df['Spent'].sum() / configs["NUM_DAYS"], 1)

    sp_spend_daily = round(df[df['Type'].str.startswith('SP')]['Spent'].sum() / configs["NUM_DAYS"], 1)
    sb_spend_daily = round(df[df['Type'].str.startswith('SB')]['Spent'].sum() / configs["NUM_DAYS"], 1)
    sd_spend_daily = round(df[df['Type'].str.startswith('SD')]['Spent'].sum() / configs["NUM_DAYS"], 1)

    sp_spend_percentage = round((sp_spend_daily / total_spend_daily) * 100, 1)
    sb_spend_percentage = round((sb_spend_daily / total_spend_daily) * 100, 1)
    sd_spend_percentage = round((sd_spend_daily / total_spend_daily) * 100, 1)

    spend_percentage = {
        "Total Daily Spend": total_spend_daily,
        "SP Daily Spend": sp_spend_daily,
        "SB Daily Spend": sb_spend_daily,
        "SD Daily Spend": sd_spend_daily,
        "SP Spend %": sp_spend_percentage,
        "SB Spend %": sb_spend_percentage,
        "SD Spend %": sd_spend_percentage,
    }

    print("Daily Spend Percentages:")
    print(tabulate(spend_percentage.items(), headers=["Parameter", "Value"], tablefmt="pretty"))

def main():
    print("Configurations:")
    print(tabulate(configs.items(), headers=["Parameter", "Value"], tablefmt="pretty"))

    df = load_and_preprocess_data('AdGroupStats.csv')
    df = calculate_metrics(df)
    df = apply_constraints(df)

    output_path = 'UpdatedAdGroupStats.xlsx'
    df.to_excel(output_path, index=False)
    
    update_bulk_file(df)
    
    calculate_spend_percentage(df)
    calculate_and_print_results(df)
    

if __name__ == "__main__":
    main()

