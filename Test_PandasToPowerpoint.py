from PandasToPowerpoint import df_to_powerpoint

import pandas as pd

df = pd.DataFrame({'District':['Hampshire', 'Dorset', 'Wiltshire', 'Worcestershire'],
				   'Population':[25000, 500000, 735298, 12653],
				   'Ratio':[1.56, 7.34, 3.67, 8.23]})

df_to_powerpoint(r"test1.pptx", df,
				  col_formatters=['', ',', '.2'], rounding=['', 3, ''])

df_to_powerpoint(r"test2.pptx", df,
				 left=1, top=1, width=10, height=15,
				 col_formatters=['', ',', '.2'], rounding=['', 3, ''])