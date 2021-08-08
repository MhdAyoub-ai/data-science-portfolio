import pandas as pd
import os


def get_data_frame_basic_stats(df, include_numeric = True):
	description = df.describe(percentiles=[0.1, 0.25, 0.5, 0.75, 0.9]
                                , include='all')
	if(include_numeric):
		IQR = description.loc['75%', :] - description.loc['25%', :]
		description.loc['IQR', :] = IQR
		description.loc['lower_range', :] = description.loc['25%', :] - (1.5 * description.loc['IQR', :])
		description.loc['upper_range', :] = description.loc['75%', :] + (1.5 * description.loc['IQR', :])
	return description

def get_value_counts(df, columns):
    column_values_dict = {}

    for col in columns:
        count_df = pd.DataFrame(df[col].value_counts().reset_index())
        count_df.columns = [col, 'count']

        percentage_df = pd.DataFrame(df[col].value_counts(normalize=True).reset_index(drop=True))
        percentage_df.columns = ['percentage']

        full_data_frame = pd.concat([count_df, percentage_df], axis=1)

        column_values_dict[col] = full_data_frame
    
    return column_values_dict

def get_duplicate_rows(df):
    duplicates = df.duplicated()
    return df[duplicates]

def get_duplicate_rows_with_sum(df):
    counted_rows_df = df.groupby(df.columns.tolist(), as_index=False).size()
    duplicate_rows_df = counted_rows_df[counted_rows_df['size'] > 1]
    return duplicate_rows_df.rename(columns={'size':'times_repeated'}).reset_index()

def get_rows_with_missing_values(df):
    missing = df.isnull().any(axis=1)
    return df[missing]

def get_columns_unique_values(df):
    columns_dict = {}
    for col in df.columns:
        columns_dict[col] = list(df[col].unique())
    
    return columns_dict

def generate_eda_basic_report(df, path = '', file_name= 'report', include_value_counts = True):
	if(path == ''):
		path = os.getcwd()
	
	basic_stats_df = get_data_frame_basic_stats(df,not include_value_counts)
	duplicate_rows_df = get_duplicate_rows_with_sum(df)
	null_containing_rows = get_rows_with_missing_values(df)
	
	if(include_value_counts):
		per_column_value_counts = get_value_counts(df, df.columns)    
    
	with pd.ExcelWriter(os.path.join(path, '{}.xlsx'.format(file_name)), engine='xlsxwriter') as writer: # pylint: disable=abstract-class-instantiated
		basic_stats_df.to_excel(writer, sheet_name='stats')
		duplicate_rows_df.to_excel(writer, sheet_name='duplicate_rows', index=False)
		null_containing_rows.to_excel(writer, sheet_name='missing values', index=False)
		if(include_value_counts):
			for key in per_column_value_counts:
				per_column_value_counts[key].to_excel(writer, sheet_name='{}_values'.format(key), index=False)
		writer.save()    


