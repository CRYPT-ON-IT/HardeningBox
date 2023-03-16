import re
from Errors import throw

class CISPdfScrapper:
    
    def __init__(self, pdf2txt, output_filepath):
        self.pdf2txt = pdf2txt
        self.output_filepath = output_filepath

    """
        This function will cut the txt to retrieve the policies only.
    """
    def LimitTxtToPoliciesOnly(self):
        recommendation_cut = self.pdf2txt.split('\nRecommendations\n')[1] # keep everything after Recommendations
        appendix_cut = recommendation_cut.split('\nAppendix: Summary Table\n')[0] # keep everything before Appendix
        self.pdf2txt = appendix_cut


    """
        This function will get the policy level from its name
    """
    def ParsePolicyName(self, policy_name):
        # Get level 
        final_level = ""
        possible_levels = ['(L1)','(L2)','(NG)', '(BL)']
        for level in possible_levels:
            if level in policy_name:
                final_level = level
        return final_level

    """
        This function will identify the order of the different paragraphs.
    """
    def setParagraphsOrder(self, policy):
        dict_index = {}

        description=False
        if 'Description:' in policy:
            dict_index['Description:'] = policy.find('Description:')
            description=True

        rationale=False
        if 'Rationale:' in policy:
            dict_index['Rationale:'] = policy.find('Rationale:')
            rationale=True


        impact=False
        if 'Impact:' in policy:
            dict_index['Impact:'] = policy.find('Impact:')
            impact=True

        audit=False
        if 'Audit:' in policy:
            dict_index['Audit:'] = policy.find('Audit:')
            audit=True

        remediation=False
        if 'Remediation:' in policy:
            dict_index['Remediation:'] = policy.find('Remediation:')
            remediation=True

        defaultvalue=False
        if 'Default Value:' in policy:
            dict_index['Default Value:'] = policy.find('Default Value:')
            defaultvalue=True

        sorted_ = list({k: v for k, v in sorted(dict_index.items(), key=lambda item: item[1])})

        return sorted_, description, rationale, impact, audit, remediation, defaultvalue


    """ 
        This function will fetch a txt file containing a CIS Benchmark PDF
        content, to retreive any information about policies (Default Value,
        Recommended Value, Impact, Description, Rationale). It also will
        transform the output to a CSV file.
    """
    def ScrapPdfData(self):
        self.LimitTxtToPoliciesOnly()
        # Transform text into a list of policies, split is based on title : "1.1.1 (L1)" with a regex
        cis_policies = re.split(r"(\d+[\.\d+]+ .*\nProfile Applicability)",self.pdf2txt)
        cis_policies.pop(0)
        cis_policies = [''.join(cis_policies[i:i+2]) for i in range(0, len(cis_policies), 2)]

        # Add csv header to csv output
        try:
            f = open(self.output_filepath, 'w+')
            f.write('"ID","Level","Policy Name","Default Value","Recommended Value","Impact","Description","Rationale","Remediation"\n')
            f.close()
        except:
            throw("Couldn't write to output filepath, please verify you have rights to write, exiting.", "highs")

        for policy in cis_policies:
            policy = re.sub(r'\d* \| P a g e', '', policy) # Remove page strings
            policy = policy.split('CIS Controls:')[0] # remove CIS Control part

            id = re.findall(r'(^\d+[\.\d]+) ', policy)[0] # Retreive policy ID
            policy_name = re.findall(id+r' (.*)', policy)[0]

            sorted_, description, rationale, impact, audit, remediation, defaultvalue = self.setParagraphsOrder(policy)

            if description:
                description_index = sorted_.index('Description:')
                if description_index >= len(sorted_)-1:
                    next_val = r'\n(.*)'
                else:
                    next_val = sorted_[description_index+1]
                
                description_content = re.findall(r'Description:\n((.|\n)*?)'+next_val, policy)
                if len(description_content) > 0:
                    description_content = description_content[0][0].replace('\n','').replace("\"","\'").encode("ascii", "ignore").decode() # Retreive description
                else:
                    description_content = ''
                
                recommended_value = re.findall(r'(?<=The recommended state for this setting is).*?(?=\.)', description_content) # Windows recommended value
                if len(recommended_value) == 0:
                    recommended_value = re.findall(r'(?=It is recommended).*?(?=\.)', description_content) # IIS recommended value
                if len(recommended_value) != 0:
                    recommended_value = recommended_value[0].replace('\n','').replace("\"","\'").encode("ascii", "ignore").decode()
                else:
                    recommended_value = ""
            else:
                description_content = ''
                recommended_value = ''

            if rationale:
                rationale_index = sorted_.index('Rationale:')
                if rationale_index >= len(sorted_)-1:
                    next_val = r'\n(.*)'
                else:
                    next_val = sorted_[rationale_index+1]
                rationale_content = re.findall(r'Rationale:\n((.|\n)*?)'+next_val, policy)

                if len(rationale_content) > 0:
                    rationale_content = rationale_content[0][0].replace('\n','').replace("\"","\'").encode("ascii", "ignore").decode() # Retreive rationale
                else:
                    rationale_content = ''
            else:
                rationale_content = ''

            if audit:
                audit_index = sorted_.index('Audit:')
                if audit_index >= len(sorted_)-1:
                    next_val = r'\n(.*)'
                else:
                    next_val = sorted_[audit_index+1]
                #audit_content = re.findall(r'Audit:\n((.|\n)*?)'+next_val, policy)[0][0].replace('\n','').replace("\"","\'") # Retreive audit

            if remediation:
                remediation_index = sorted_.index('Remediation:')
                if remediation_index >= len(sorted_)-1:
                    next_val = r'\n(.*)'
                else:
                    next_val = sorted_[remediation_index+1]
                out =[]
                strings = policy.splitlines()
                for index, line in enumerate(strings):
                    if 'Computer Configuration\\' in line or 'User Configuration\\' in line:
                        if index+1 < len(strings) and '\\' in strings[index+1]:
                            line+=strings[index+1]
                        out.append(line.strip())
                remediation_content = ';'.join(out)

            if impact:
                impact_index = sorted_.index('Impact:')
                if impact_index >= len(sorted_)-1:
                    next_val = r'\n(.*)'
                else:
                    next_val = sorted_[impact_index+1]

                impact_content = re.findall(r'Impact:\n((.|\n)*?)'+next_val, policy)
                if len(impact_content) > 0:
                    impact_content = impact_content[0][0].replace('\n','').replace("\"","\'").encode("ascii", "ignore").decode() # Retreive impact
                else:
                    impact_content = ''
            else:
                impact_content = ''

            if defaultvalue:
                defaultvalue_index = sorted_.index('Default Value:')
                if defaultvalue_index >= len(sorted_)-1:
                    next_val = r'\n(.*)'
                else:
                    next_val = sorted_[defaultvalue_index+1]
                defaultvalue_content = re.findall(r'Default Value:\n((.|\n)*?)'+next_val, policy)
                if len(defaultvalue_content) > 0:
                    defaultvalue_content = defaultvalue_content[0][0].replace('\n','').replace("\"","\'").encode("ascii", "ignore").decode() # Retreive default value
                else:
                    defaultvalue_content = ''
            else:
                defaultvalue_content = ''

            # parse policy name
            level = self.ParsePolicyName(policy_name)

            f = open(self.output_filepath, 'a')
            f.write('"'+id+'","'+level+'","'+policy_name+'","'+defaultvalue_content+'","'+recommended_value+'","'+impact_content+'","'+description_content+'","'+rationale_content+'","'+remediation_content+'"\n')
            f.close()

        
