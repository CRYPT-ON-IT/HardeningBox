import re
from Errors import throw

class CISPdfScrapper:
    
    def __init__(self, pdf2txt, output_filepath):
        self.pdf2txt = pdf2txt
        self.output_filepath = output_filepath

    """ 
        This function will fetch a txt file containing a CIS Benchmark PDF
        content, to retreive any information about policies (Default Value,
        Recommended Value, Impact, Description, Rationale). It also will
        transform the output to a CSV file.
    """
    def ScrapPdfData(self):
        # Transform text into a list of policies, split is based on title : "1.1.1 (L1)" with a regex
        cis_policies = re.split(r"(\d+[\.\d+]+ \(L[1-3]\))",self.pdf2txt)
        cis_policies.pop(0)
        cis_policies = [''.join(cis_policies[i:i+2]) for i in range(0, len(cis_policies), 2)]

        # Add csv header to csv output
        try:
            f = open(self.output_filepath, 'w+')
            f.write('"ID","Default Value","Recommended Value","Impact","Description","Rationale"\n')
            f.close()
        except:
            throw("Couldn't write to output filepath, please verify you have rights to write, exiting.", "highs")

        for policy in cis_policies:
            policy = re.sub(r'\d* \| P a g e', '', policy) # Remove page strings

            id = re.findall(r'^\d+[\.\d]+ \(L[1-3]\)', policy)[0][:-5] # Rereive policy ID

            try:
                default_value = re.findall(r'Default Value:\n(.*)', policy)[0].replace("\"","\'").replace('\n','') # Retreive default value
            except:
                default_value = ''

            try:
                recommended_value = re.findall(r'(?<=The recommended state for this setting is).*', policy)[0].replace('\n','').replace("\"","\'").replace('Rationale:','') # Retreive recommended value
                recommended_value = re.sub(r'^.+?(?=[0-9a-zA-Z])', '', recommended_value)
            except:
                recommended_value = ''

            try:
                impact = re.findall(r'Impact:\n((.|\n)*?)Audit:', policy)[0][0].replace('\n', '').replace("\"","\'") # Retreive impact
            except:
                impact = ''

            try:
                description = re.findall(r'Description:\n((.|\n)*?)The recommended state for this setting is', policy)[0][0].replace('\n','').replace("\"","\'") # Retreive description
            except:
                description = ''

            try:    
                rationale = re.findall(r'Rationale:\n((.|\n)*?)Impact:', policy)[0][0].replace('\n','').replace("\"","\'") # Retreive rationale
            except:
                rationale = ''

            f = open(self.output_filepath, 'a')
            f.write('"'+id+'","'+default_value+'","'+recommended_value+'","'+impact+'","'+description+'","'+rationale+'"\n')
            f.close()