# GobyParser

解析Goby扫描结果转换为JSON并录入到MYSQL数据库的工具

## Install

```cmd
python3 -m pip install -r requirements.txt
```

## Usage

使用前需先修改代码中的数据库连接参数

解析单个Goby扫描结果
```cmd
python3 GobyParser.py filename.xlsx
```

批量解析Goby扫描结果
```cmd
python3 GobyParser.py -d dir_name
```
