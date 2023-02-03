''' Code for Simone Troxler (01.02.2023)
Merging a mapping table with virus analysis files depending on the date it was performed
Analysis in 96-well plates (Position variable A01 - H12)'''

import os
import pandas as pd
from pathlib import Path
from glob import glob

# import data from the mapping file
def import_mapping_file(mapping_path):
    return pd.read_excel(mapping_path)

# function that merge mapping file and virus analysis files
def merge_files(mapping_file, current_dir):
    files = os.listdir(current_dir)
    new_df = pd.DataFrame()
    for file in files:
        if '.xls' in file and 'Mapping' not in file:
            col_names = ['Sample', 'Patient_id', 'Position', 'Name', 'Type', 'SC2-S/N', 'C(t)_SC2-S/N',
                         'SC2-RdRP', 'C(t)_SC2-RdRP', 'PIV', 'C(t)_PIV', 'Flu_B', 'C(t)_Flu_B',
                         'AdV', 'C(t)_AdV', 'Flu_A', 'C(t)_Flu_A', 'HRV', 'C(t)_HRV', 'MPV',
                         'C(t)_MPV', 'RSV', 'C(t)_RSV', 'IC', 'C(t)_IC']

            df_current_file = pd.read_excel(os.path.join(current_dir,file),
                                            header=None, names=col_names, usecols='A:Y')
            df_current_file.drop(index=df_current_file.index[:2], axis=0, inplace=True)
            df_current_file['Laufdatum'] = file[:-4]
            mapping_file_onePlate = mapping_file.merge(df_current_file, on=['Position', 'Laufdatum'], how='inner')
            new_df = new_df.append(mapping_file_onePlate)
    return new_df

#function that replace - and + signs by 0 and 1 numbers
def replace_zerosOnes(df):
    replaced_df = df.replace(['-', '+'], ['0', '1'])
    return replaced_df


#write and save an excel file with the merged results (directory and name of the new file is in the path)
def write_excel(df,path, fileName):
    output_name = Path(fileName + '.xlsx')
    i = glob(output_name.stem + "_[0-9]*" + output_name.suffix)
    new_output_name = f"{output_name.stem}_{len(i)}{output_name.suffix}"
    df.to_excel((path + '\\' + new_output_name), index=False)


if __name__ == '__main__':

    # directory path of the folder with mapping file and virus analysis files
    current_dir = r'\\filer300\USERS3007\I0337516\Desktop\testSimone'

    # name and save path for the output excel merged file
    fileName = 'mergedMapping'
    savePath = r'\\filer300\USERS3007\I0337516\Desktop\testSimone'

    mapping_file = import_mapping_file(os.path.join(current_dir, r'Mapping_file.xlsx'))
    mapping_file.drop(columns=mapping_file.columns[-1],  axis=1,  inplace=True)
    merged_df = merge_files(mapping_file, current_dir)
    final_merged_df = replace_zerosOnes(merged_df)
    write_excel(final_merged_df, savePath, fileName)
