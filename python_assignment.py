import pandas as pd  # pip install pandas
from tabulate import tabulate  # pip install tabulate

df1 = pd.read_excel('Python Assignment.xlsx', sheet_name=1, header=10, nrows=21, usecols='D:G')
df1.drop('S No', axis='columns', inplace=True)
df1['Team Name'] = df1['Team Name'].str.replace('tech', 'Tech', case=True)

df2 = pd.read_excel('Python Assignment.xlsx', sheet_name=2, header=7, skiprows=0, nrows=21, usecols='C:G')
df2.rename(columns={'name': 'Name'}, inplace=True)

merged_df = pd.merge(df1, df2, on='Name', how='inner')
grouped_df = merged_df.groupby('Team Name')[['total_statements', 'total_reasons']].mean()

grouped_df['sum'] = (grouped_df['total_reasons'] + grouped_df['total_statements'])

grouped_df['total_statements'] = round(grouped_df['total_statements'], 2)
grouped_df['total_reasons'] = round(grouped_df['total_reasons'], 2)
grouped_df.sort_values('sum', ascending=False, inplace=True)
grouped_df.reset_index(inplace=True)
grouped_df['Team Rank'] = range(1, len(grouped_df) + 1)
grouped_df.set_index('Team Rank', inplace=True)
grouped_df.index.name = 'Team Rank'
grouped_df.drop('sum', axis='columns', inplace=True)
grouped_df = grouped_df.rename(columns={'total_statements': 'Average Statements', 'total_reasons': 'Average Reasons'})

df2['sum'] = (df2['total_reasons'] + df2['total_statements'])
df2['name_lower'] = df2['Name'].str.lower()
df2.sort_values(['sum', 'name_lower', 'total_statements'], ascending=[False, True, False], inplace=True)
df2.drop('S No', axis='columns', inplace=True)

df2['Team Rank'] = range(1, len(df2) + 1)
df2.set_index('Team Rank', inplace=True)
df2.index.name = 'Team Rank'

df2.drop(['sum', 'name_lower'], axis='columns', inplace=True)

df2 = df2.rename(columns={'total_statements': 'No. of Statements', 'uid': 'UID', 'total_reasons': 'No. of Reasons'})


print(tabulate(grouped_df, headers='keys', tablefmt='psql'), '\n')
print(tabulate(df2, headers='keys', tablefmt='psql'))
