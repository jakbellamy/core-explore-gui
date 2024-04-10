import pandas as pd

agg_by = lambda df, col: df.groupby(col).agg('sum')[['Closed Sales Volume YTD']].sort_values(by='Closed Sales Volume YTD', ascending=False)

def clean_currency(x):
    """
    This function is used to clean the currency data.
    """
    if isinstance(x, str):
        return float(x.replace('$', '').replace(',', ''))
    else:
        return x

def get_vendor_ranks(file_path):
    df = pd.read_csv("./processed_cores/" + file_path)
    bsnc_mask = df['Ven'] == 'M'
    df = df[bsnc_mask]
    df['Closed Sales Volume YTD'] = df['Closed Sales Volume YTD'].apply(clean_currency).astype('float')
    agg_by(df, 'Vendor Name').to_clipboard()

def get_agent_ranks(file_path):
    df = pd.read_csv("./processed_cores/" + file_path)
    bsnc_mask = df['Ven'] == 'M'
    df = df[bsnc_mask]
    df['Closed Sales Volume YTD'] = df['Closed Sales Volume YTD'].apply(clean_currency).astype('float')
    agg_by(df, 'Associate Name').to_clipboard()

def re_agg_file(file_path):
    """
    This function is used to re-aggregate the file.
    The idea is the user will make the changes then copy the new file to the clipboard.
    The file is then read in and re-aggregated and saved to ./processed_cores/{file_name} - UPDATED.csv
    """
    df = pd.read_clipboard()
    bsnc_mask = df['Ven'] == 'M'
    df = df[bsnc_mask]
    df['Closed Sales Volume YTD'] = df['Closed Sales Volume YTD'].apply(clean_currency).astype('float')
    df['Supreme Volume'] = df.apply(lambda row: row['Closed Sales Volume YTD'] if row['Vendor Name'] == 'Supreme Lending' else 0, axis=1)
    df = df.groupby('Associate Name').agg('sum')[['Closed Sales Volume YTD', 'Supreme Volume']]
    df['Supreme Volume %'] = df['Supreme Volume'] / df['Closed Sales Volume YTD']
    df = df.sort_values(by='Supreme Volume', ascending=False).head(6)