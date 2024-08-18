import pandas as pd

data = {"idchainТТ": [14],	"namechainТТ": ["АБВ"],	"yymm":	[202304], "tm":	"Aquafresh", "count_pos_by_chain_yymm_tm":	[9],
        "CIBnoVATadmin_by_brand":	[0], 	"CIBnoVATadmin_USTM":	[0], "CIBnoVATadmin_promo":	[0],
        "tt_CIBnoVATadmin_USTM": [2],	"tt_CIBnoVATadmin_promo": [0],	"Baseline_rub": [1],
      "coef_promo_USTM_rub": [1],	"coef_promo_promo_rub":	[5], 	"offtake_forcast_rub":	[6],
        "Baseline_pcs": [2],	"coef_promo_USTM_pcs": [4],
      "coef_promo_promo_pcs": [1],	"offtake_forcast_pcs": [3],	"coef_corel_pcs": [10],	"coef_corel_rub":	[5],
      "Baseline_so_rub":	[3], 	"Coef_offtake_for_so_forecast_rub": [4],	"so_forcast_rub":	[3], 	"Baseline_so_pcs":	[2],
      "Coef_offtake_for_so_forecast_pcs":	[2], 	"so_forcast_pcs":	[1], 	"model_version": [1], 	"max_fact_sales_date":	["5-31-2024"]}

df = pd.DataFrame(data)
def write_row_excel(filename: str, df: pd.DataFrame, sheetname: str) -> None:
    """Write data to an existing Excel file"""
    try:
        with pd.ExcelWriter(filename,
                            mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheetname, index=True)
    except PermissionError:
        print("Close the file in Excel and try again.")
    except FileNotFoundError:
        print(f'File {filename} not found')

def update_excel(filename: str, sheetname: str, df:pd.DataFrame, idchainТТ:tuple, yymm:tuple,
                     tm:str, model_version:tuple, max_fact_sales_date:str) -> None:
    """Filter and deleting rows from an existing Excel file"""
    old_df = pd.read_excel(filename, sheet_name=sheetname, index_col=0)
    new_df = old_df[~(old_df['idchainТТ'].isin(idchainТТ))
                    & ~(old_df['yymm'].isin(yymm))
                    & ~(old_df['tm'].str.contains(tm, na=False))
                    & ~(old_df['model_version'].isin(model_version))
                    & ~(old_df['max_fact_sales_date'].str.contains(max_fact_sales_date, na=False))]
    df_to_file = pd.concat([new_df, df])
    print(df_to_file)
    write_row_excel(filename, df_to_file, sheetname)

update_excel("main_df.xlsx", "Sheet1", df,(14, 916,),
                 (202304,), "Aquafresh", (1,), "5-31-2023")