import requests

class UpdateMainCsv():

    def __init__(self, original_dataframe, original_filepath, adding_dataframe = None, adding_filepath = None):
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

    def AddMicrosoftLinks(self):
        self.original_dataframe.insert(10, 'MicrosoftLink', None)
        self.original_dataframe.insert(11, 'PossibleValues', None)

        for index, policy in self.original_dataframe.iterrows():
            name = policy['Name']
            lower_name = name.lower()
            policy_name = lower_name.replace(' ', '-').replace(':', '').replace('(','').replace(')','')

            if policy['Name'].startswith('Device Guard: '):
                # There's only one doc for Device Guard
                full_link = 'https://docs.microsoft.com/en-us/windows/security/identity-protection/credential-guard/credential-guard-manage'
                self.original_dataframe.at[index, 'MicrosoftLink'] = full_link
                continue

            elif policy['Name'].startswith('Device Installation: '):
                full_link = 'https://docs.microsoft.com/en-us/windows/client-management/manage-device-installation-with-group-policy'
                self.original_dataframe.at[index, 'MicrosoftLink'] = full_link
                continue

            elif policy['Category'] in ['System Services','Administrative Templates: LAPS', 'Administrative Templates: Control Panel', 'MSS (Legacy)', 
            'Administrative Templates: Network', 'Administrative Templates: Start Menu and Taskbar', 'Administrative Templates: System', 
            'Administrative Templates: Windows Components', 'Microsoft Defender Antivirus', 'Microsoft Edge', 'PowerShell']:
                # There's no Microsoft link for this policiy category
                continue

            elif policy['Category'] == 'Windows Firewall':
                # There's only one doc for firewall configuration
                full_link = 'https://docs.microsoft.com/en-us/windows/security/threat-protection/windows-firewall/best-practices-configuring'
                self.original_dataframe.at[index, 'MicrosoftLink'] = full_link
                continue

            elif policy['Category'] == 'MS Security Guide':
                full_link = 'https://docs.microsoft.com/en-us/windows/security/threat-protection/windows-security-configuration-framework/windows-security-baselines'
                self.original_dataframe.at[index, 'MicrosoftLink'] = full_link
                continue

            elif policy['Category'] == 'Microsoft Defender Application Guard':
                full_link = 'https://docs.microsoft.com/en-us/windows/security/threat-protection/microsoft-defender-application-guard/configure-md-app-guard#application-specific-settings'
                self.original_dataframe.at[index, 'MicrosoftLink'] = full_link
                continue

            elif policy['Category'] == 'Microsoft Defender Exploit Guard':
                microsoft_link = "https://docs.microsoft.com/en-us/microsoft-365/security/defender-endpoint/attack-surface-reduction-rules-reference?view=o365-worldwide#"
                policy_name = policy_name.replace('asr-','')
                r = requests.get(microsoft_link + policy_name)
                if r.status_code == 200:
                    full_link = microsoft_link + policy_name
                    self.original_dataframe.at[index, 'MicrosoftLink'] = full_link
                elif r.status_code not in [200, 404]:
                    print('An error occured with unexpected status code ' + str(r.status_code))
                continue

            elif policy['Category'] == 'Advanced Audit Policy Configuration':
                microsoft_link = 'https://docs.microsoft.com/en-us/windows/security/threat-protection/auditing/audit-'
                r = requests.get(microsoft_link + policy_name)
                if r.status_code == 200:
                    full_link = microsoft_link + policy_name
                    self.original_dataframe.at[index, 'MicrosoftLink'] = full_link
                elif r.status_code not in [200, 404]:
                    print('An error occured with unexpected status code ' + str(r.status_code))
                continue

            elif policy['Category'] in ['Account Policies', 'User Rights Assignment', 'Security Options']:
                microsoft_link = 'https://docs.microsoft.com/en-us/windows/security/threat-protection/security-policy-settings/'
                r = requests.get(microsoft_link + policy_name)
                if r.status_code == 200:
                    full_link = microsoft_link + policy_name
                    self.original_dataframe.at[index, 'MicrosoftLink'] = full_link

                    # Retreive possible values
                    response_list = r.text.splitlines()
                    line_number = 0
                    for line in response_list:
                        if line == '<h3 id="possible-values">Possible values</h3>':
                            break
                        line_number+=1

                    if response_list[line_number+1] == "<ul>":
                        possible_values = []
                        while not response_list[line_number+2].startswith("</ul>"):
                            possible_values.append(response_list[line_number+2].replace('<li>','').replace('</li>','').replace('<p>','').replace('</p>','').replace('<em>','').replace('</em>','').replace('<strong>','').replace('</strong>',''))
                            line_number+=1

                    self.original_dataframe.at[index, 'PossibleValues'] = possible_values
                elif r.status_code not in [200, 404]:
                    print('An error occured with unexpected status code ' + str(r.status_code))


            self.original_dataframe.to_csv('linked_' + self.original_filepath)