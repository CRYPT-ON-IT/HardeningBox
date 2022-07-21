class UpdateMainCsv():

    def __init__(self, original_dataframe, original_filepath, adding_dataframe, adding_filepath):
        self.original_dataframe = original_dataframe
        self.original_filepath = original_filepath
        self.adding_dataframe = adding_dataframe
        self.adding_filepath = adding_filepath

    def AddAuditResult(self):
        output_column = input('What will be the name of the output column (e.g. : context1) : ')
        if output_column == '':
            print('No output column provided.')

        self.original_dataframe.insert(15, output_column, None)

        for index, policy in self.adding_dataframe.iterrows():
            policy_name = policy['Name']
            audit_result = policy['Result']

            self.original_dataframe.loc[self.original_dataframe['Name'] == policy_name, output_column] = audit_result
            self.original_dataframe.to_csv('compare_' + self.original_filepath, index=False)