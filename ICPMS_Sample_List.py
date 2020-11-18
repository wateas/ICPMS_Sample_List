import re
import xlwings as xw
import pandas as pd

@xw.sub  # only required if you want to import it or run it via UDF Server
def main():

    wb = xw.Book.caller()
    sht1 = wb.sheets['Import']
    sht2 = wb.sheets['ElementsBySample']
    
    df = import_sample_info(sht1)
    
    df_processed = Elements_By_Sample(df)
    
    export_table(sht2, df_processed)
    
    
def import_sample_info(sht):
    """
    Connects with workbook worksheet "Import"; imports sample information table
    as pandas DataFrame.
    
    """        
    # https://stackoverflow.com/questions/34392805/a-whole-sheet-into-a-panda-dataframe-with-xlwings
    df = sht.range('A1').options(pd.DataFrame, 
                              header=1,
                              index=False, 
                              expand='table').value
    
    return df
    

def Elements_By_Sample(df):
    '''
    Parameters
    ----------
    df : Dataframe of Sample Information and Required analysis.  Row for each -
    unique analysis code / sample, so can be multiple rows for the sample 
    sample.

    Returns
    -------
    Dataframe with one row per unique sample, with a list of needed analytes -
    (just element names, not LIMS analysis codes).
    
    '''
    # Remove unwanted analyses from sample info list.
    unwanted = ["MET_DIG", "HG_CV", "HG_CV_SL", "HG_DIG", "DRYWT", 
                    "SLDG_WT", "SLG_WT_HG", "SLG_WT_H"]
    df = df[~df["Analysis Code"].isin(unwanted)]
    
    # Create new column "Analyte" from column "Analysis Code".
    regex_pat = re.compile(r'_ICPMS|_DRYWT|_SL|_AQ|_DW|_CV')
    df['Analyte'] = df['Analysis Code'].str.replace(regex_pat, '')
    
    # Need to remove duplicates.  Must get rid of column "Analysis Code" prior.
    df = df.drop(columns=['Analysis Code', 'Location Code'])
    # Must fill in column "Location Code 2" with placeholder value so that -
    # function pd.pivot_table() doesn't filter out values.
    df["Location Code 2"].fillna(value="-", inplace=True)
    # Use argument "ignore_index" to create new, consecutive index values.
    df = df.drop_duplicates(subset=None, keep='first', ignore_index=True)
    
    # Use pivot table to reorganize table so that there is one row per unique -
    # sample and a corresponding list of elements.
    # Note that the table will be sorted by order passed for parameter "index".
    # If a row is NaN for a particular index, the entire row will be omitted.
    df = df.pivot_table(index=['LIMS #', 'Sample Location', 
                                'Collection Date', 'Location Code 2'],
                            values='Analyte',
                            aggfunc=lambda x: ', '.join(x))
    
    return df


def export_table(sht, df):
    """
    Clears existing information and prints processed sample information table 
    on excel workbook worksheet "ElementsBySample".
    """
    
    sht.range('A1').expand('table').clear_contents()
    
    sht.range('A1').value = df
    

if __name__ == "__main__":
    xw.Book("ICPMS_Sample_List.xlsm").set_mock_caller()
    main()
    
