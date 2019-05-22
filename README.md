# WebBaoBiao
Web Baobiao App for Trading Finance Department

1. 上传报表模板，进数据库Table各张报表
2. 拆分到数据库，进数据库Table各位用户
3. 生成报表
4. 下载报表

## Environment

Use ``pip install -r ./requirements.txt`` to install the dependencies.


## Demand
1.同一用户在同一报表的同一位置填写多个数据（重新考虑每张user表的主键）  
2.增加保存分析型报表的模块  
3.校验时，校验每位用户填写的数据，而非最终汇总的。 需要修改每张user表
4.模板中支持四则运算，最好带括号，重新考虑下模板中填写的字符串和拆分方法  
（按操作符分割，存操作数序列和操作符序列，取值后表达式求值。） 大部分完成
