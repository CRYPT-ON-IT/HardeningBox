import pandas as pd

# Based on csv file with names of policy, it will retreive entire policies corresponding

base_file = '/Users/guillaumederybel/Documents/githubprojects/HardeningBox/sonepar_server/name.csv'
context_file_1 = '/Users/guillaumederybel/Documents/githubprojects/HardeningBox/GetFullPoliciesByName/full.csv'

base_df = pd.read_csv(base_file, encoding='latin1')
context_df_1 = pd.read_csv(context_file_1, encoding='latin1')
#print(context_df_1.columns.values.tolist())
output_df = pd.DataFrame(columns=context_df_1.columns.values.tolist())

for index, policy in base_df.iterrows():
    search = context_df_1.loc[context_df_1['Name'] == policy['Name']].values.tolist()   
    if len(search)>0:
        output_df.loc[len(output_df)] = search[0]
    else:
        print("The following policy was not found : " + policy['Name'])

output_df.to_csv("GetFullPoliciesByName/output_full.csv", index=False)