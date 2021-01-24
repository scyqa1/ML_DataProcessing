import pandas as pd
import numpy as np
import pingouin as pg


io = r'Lake_v2.xls'

data = pd.read_excel(io, sheet_name=2, usecols=[2,3,4])
data.head()
print(len(data))
for i in range(len(data)):
    print(data.loc[i])

corr=pg.pairwise_corr(data, method='pearson')
spearman_corr=pg.pairwise_corr(data, method='spearman')
kendall_corr=pg.pairwise_corr(data, method='kendall')
bicor_corr=pg.pairwise_corr(data, method='bicor')
skipped_corr=pg.pairwise_corr(data, method='skipped')
corr=corr.append(spearman_corr)
corr=corr.append(kendall_corr)
corr=corr.append(bicor_corr)
corr=corr.append(skipped_corr)

corr.to_excel('correlation.xls')
