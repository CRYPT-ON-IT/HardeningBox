import requests
import pandas as pd
from Errors import throw

class UpdateMainCsv():
    """
        This class will updatet csv
        by adding diffrent columns.
    """

    def __init__(self, original_dataframe, original_filepath,
     adding_dataframe = None, adding_filepath = None, output_filepath = None):
        self.original_dataframe = original_dataframe
        self.original_filepath = original_filepath
        self.adding_dataframe = adding_dataframe
        self.adding_filepath = adding_filepath
        self.output_filepath = output_filepath

    def add_audit_result(self):
        """
            This function will add a new column
            from a csv to another by using pandas.
        """
        max_index = len(self.original_dataframe.columns)

        output_column_name = input("""
        What will be the name of the output column (e.g. : context1) : 
        """)
        if output_column_name == '':
            throw('No output column provided, exiting.', 'high')

        output_column_index = input("""
        What will be the index of the output column (max : """ + str(max_index) + """) :
        """)
        try:
            output_column_index = int(output_column_index)
        except ValueError:
            throw('Output index is not an integers, exiting.', 'high')
        if 0 < output_column_index > max_index:
            throw('Output index out of range, exiting.', 'high')

        self.original_dataframe.insert(output_column_index, output_column_name, None)

        if self.output_filepath == "":
            self.output_filepath = input('How should I name the output CSV ? : ')

        for _, policy in self.adding_dataframe.iterrows():
            policy_name = policy['Name']
            audit_result = policy['Result']

            self.original_dataframe.loc[
                self.original_dataframe['Name'] == policy_name, output_column_name
            ] = audit_result
        try:
            self.original_dataframe.to_csv(self.output_filepath, index=False)
        except:
            throw(
                "Couldn't create CSV file, please check you have rights\
 to wright in this folder, exiting.",
                "high"
            )

    def add_microsoft_links(self):
        """
            This function will add a new column
            to an hardening file with some Microsoft
            Links to help the user.
        """
        self.original_dataframe = self.original_dataframe.assign(MicrosoftLink=None)
        self.original_dataframe = self.original_dataframe.assign(PossibleValues=None)

        print('\033[93mFetching Microsoft website, it might take less than a minute...\n\033[0m')
        for index, policy in self.original_dataframe.iterrows():
            name = policy['Name']
            lower_name = name.lower()
            policy_name = lower_name.replace(
                ' ', '-'
                ).replace(':', ''
                          ).replace('(',''
                                    ).replace(')','')

            if policy['Name'].startswith('Device Guard: '):
                # There's only one doc for Device Guard
                full_link = 'https://docs.microsoft.com/en-us/windows/security/identity\
-protection/credential-guard/credential-guard-manage'
                self.original_dataframe.at[index, 'MicrosoftLink'] = full_link
                continue

            elif policy['Name'].startswith('Device Installation: '):
                full_link = 'https://docs.microsoft.com/en-us/windows/client-management\
/manage-device-installation-with-group-policy'
                self.original_dataframe.at[index, 'MicrosoftLink'] = full_link
                continue

            elif policy['Category'] in ['System Services','Administrative Templates: LAPS',
            'Administrative Templates: Control Panel', 'MSS (Legacy)', 
            'Administrative Templates: Network', 'Administrative Templates: Start Menu and Taskbar',
            'Administrative Templates: System', 'Administrative Templates: Windows Components',
            'Microsoft Defender Antivirus', 'Microsoft Edge', 'PowerShell']:
                # There's no Microsoft link for this policiy category
                continue

            elif policy['Category'] == 'Windows Firewall':
                # There's only one doc for firewall configuration
                full_link = 'https://docs.microsoft.com/en-us/windows/security/\
threat-protection/windows-firewall/best-practices-configuring'
                self.original_dataframe.at[index, 'MicrosoftLink'] = full_link
                continue

            elif policy['Category'] == 'MS Security Guide':
                full_link = 'https://docs.microsoft.com/en-us/windows/security/\
threat-protection/windows-security-configuration-framework/windows-security-baselines'
                self.original_dataframe.at[index, 'MicrosoftLink'] = full_link
                continue

            elif policy['Category'] == 'Microsoft Defender Application Guard':
                full_link = 'https://docs.microsoft.com/en-us/windows/security/\
threat-protection/microsoft-defender-application-guard/configure-md-app-guard\
#application-specific-settings'
                self.original_dataframe.at[index, 'MicrosoftLink'] = full_link
                continue

            elif policy['Category'] == 'Microsoft Defender Exploit Guard':
                microsoft_link = "https://docs.microsoft.com/en-us/microsoft-365/security/\
defender-endpoint/attack-surface-reduction-rules-reference?view=o365-worldwide#"
                policy_name = policy_name.replace('asr-','')
                response = requests.get(microsoft_link + policy_name, timeout=5)
                if response.status_code == 200:
                    full_link = microsoft_link + policy_name
                    self.original_dataframe.at[index, 'MicrosoftLink'] = full_link
                elif response.status_code not in [200, 404]:
                    print(f'An error occured with unexpected status code {response.status_code}')
                continue

            elif policy['Category'] == 'Advanced Audit Policy Configuration':
                microsoft_link = 'https://docs.microsoft.com/en-us/windows/security/\
threat-protection/auditing/audit-'
                response = requests.get(microsoft_link + policy_name, timeout=5)
                if response.status_code == 200:
                    full_link = microsoft_link + policy_name
                    self.original_dataframe.at[index, 'MicrosoftLink'] = full_link
                elif response.status_code not in [200, 404]:
                    throw(
                        f'An error occured with unexpected status code {response.status_code}',
                        'high'
                    )
                continue

            elif policy['Category'] in [
                'Account Policies',
                'User Rights Assignment',
                'Security Options'
            ]:
                microsoft_link = 'https://docs.microsoft.com/en-us/windows/security/\
threat-protection/security-policy-settings/'
                response = requests.get(microsoft_link + policy_name, timeout=5)
                if response.status_code == 200:
                    full_link = microsoft_link + policy_name
                    self.original_dataframe.at[index, 'MicrosoftLink'] = full_link

                    # Retreive possible values
                    response_list = response.text.splitlines()
                    line_number = 0
                    for line in response_list:
                        if line == '<h3 id="possible-values">Possible values</h3>':
                            break
                        line_number+=1

                    if (len(response_list) >= line_number+1
                        and response_list[line_number+1] == "<ul>"):
                        possible_values = []
                        while not response_list[line_number+2].startswith("</ul>"):
                            possible_values.append(response_list[line_number+2].replace('<li>',''
                                ).replace('</li>',''
                                ).replace('<p>',''
                                ).replace('</p>',''
                                ).replace('<em>',''
                                ).replace('</em>',''
                                ).replace('<strong>',''
                                ).replace('</strong>',''))
                            line_number+=1

                    self.original_dataframe.at[index, 'PossibleValues'] = possible_values
                elif response.status_code not in [200, 404]:
                    throw(
                        f'An error occured with unexpected status code {response.status_code}',
                        "high"
                    )

        output_filepath = input('How should we name the output file ? : ')
        try:
            self.original_dataframe.to_csv(output_filepath, index=False)
        except:
            throw(
                "Couldn't create output file, verify you have rights\
to write in this folder, exiting.",
                "high"
            )

    def add_scrapped_data_to_csv(self):
        """
            This function will add scrapped data
            from a CIS Benchmark PDF to a CSV file.
        """
        self.original_dataframe = self.original_dataframe.assign(Impact=None)
        self.original_dataframe = self.original_dataframe.assign(Rationale=None)
        self.original_dataframe = self.original_dataframe.assign(Description=None)
        self.original_dataframe = self.original_dataframe.assign(ScrappedDefaultValue=None)
        self.original_dataframe = self.original_dataframe.assign(ScrappedRecommendedValue=None)
        self.original_dataframe = self.original_dataframe.assign(Remediation=None)
        self.original_dataframe = self.original_dataframe.assign(Level=None)

        for _, policy in self.original_dataframe.iterrows():
            search = self.adding_dataframe.loc[self.adding_dataframe['ID'] == policy['ID']]
            # Checking ID
            if search['Level'].values.size == 0:
                id_1 = policy['ID']
                id_1 = id_1.split('.')
                id_1.pop()
                id_1 = '.'.join(id_1)
                search = self.adding_dataframe.loc[self.adding_dataframe['ID'] == id_1]

            # Checking ID 2
            if search['Level'].values.size == 0:
                id_2 = id_1
                id_2 = id_2.split('.')
                id_2.pop()
                id_2 = '.'.join(id_2)
                search = self.adding_dataframe.loc[self.adding_dataframe['ID'] == id_2]

            if search['Level'].values.size == 0:
                print('\033[93mWarning: Unable to get data from ' +
                      policy['ID'] +
                      ' in scrapped content.\033[0m\n'
                )

            search_impact = search['Impact'].values
            if search_impact.size > 0:
                policy['Impact'] = search_impact[0]

            search_description = search['Description'].values
            if search_description.size > 0:
                policy['Description'] = search_description[0]

            search_rationale = search['Rationale'].values
            if search_rationale.size > 0:
                policy['Rationale'] = search_rationale[0]

            search_recommended_value = search['Recommended Value'].values
            if search_recommended_value.size > 0:
                policy['ScrappedRecommendedValue'] = search_recommended_value[0]

            search_default_value = search['Default Value'].values
            if search_default_value.size > 0:
                policy['ScrappedDefaultValue'] = search_default_value[0]

            search_remediation = search['Remediation'].values
            if search_remediation.size > 0:
                policy['Remediation'] = search_remediation[0]

            search_level = search['Level'].values
            if search_level.size > 0:
                policy['Level'] = search_level[0]

        try:
            self.original_dataframe.to_csv(self.output_filepath, index=False)
        except:
            throw("Couldn't create output file, verify you have\
 rights to write in this folder, exiting.",
                  "high"
            )

    def merge_two_csv(self):
        """
            This function will merge two csv files
            by adding diffrent policies.
        """

        first_dataframe = self.original_dataframe
        second_dataframe = self.adding_dataframe

        frames = [first_dataframe, second_dataframe]
        new_dataframe = pd.concat(frames)
        count1 = len(new_dataframe.index)
        # we should to keep policy with defined level
        new_dataframe = new_dataframe.drop_duplicates(subset=["Name"], keep='first')
        new_dataframe = new_dataframe.sort_values("Category")
        count2 = len(new_dataframe.index)

        print("Total of policies : ",count1, '--->', "reduced to : ", count2)
        try:
            new_dataframe.to_csv(self.output_filepath, index=False)
        except:
            throw("Couldn't create output file, verify you have rights \
to write in this folder, exiting.",
                  "high"
            )
