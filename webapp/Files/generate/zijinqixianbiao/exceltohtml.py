import codecs
import pandas as pd
xd = pd.ExcelFile('zijinqixianbiao_2019_03_31.xlsx')
pd.set_option('display.max_colwidth',1000)#设置列的宽度，以防止出现省略号
df = xd.parse()
with codecs.open('zijinqixianbiao_2019_03_31.html','w') as html_file:
    html_file.write(df.to_html(header = True,index = False))