from pyPreservica import *
import pandas as pd
import secret

content = ContentAPI(username=secret.username,password=secret.password,server="unilever.preservica.com")

file = r"C:\Users\Chris.Prince\OneDrive - Unilever\Quickfix_Reference.xlsx"

df = pd.read_excel(file)

reflist = df['Reference'].values.tolist()

metadata_filters = {"rm"}
for r in reflist:
    content.search_index_filter_list(query=r,filter_values=)