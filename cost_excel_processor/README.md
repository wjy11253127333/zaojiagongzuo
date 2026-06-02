# 造价Excel清单自动化处理程序

> 专为建筑施工企业解决Excel清单不规范导致成本分析难的问题

## 核心功能

### 1. 智能解析与分类
- 自动识别并读取Excel工作表数量及结构（单表/多表）
- 多级合并表头自动拆分为标准一级表头
- 支持 `.xlsx` / `.xls` 格式

### 2. 数据结构化拆解
- 必填字段校验：`业态、项目名称、项目特征、计量单位、工程量、综合单价、综合合价、备注`
- 缺失字段自动提示（交互模式可手动映射）
- 建立量价关联模型，输出标准化数据

### 3. 高兼容性上传
- 支持多格式Excel导入
- 异常数据提示
- 多工作表自动整合（纵向合并 + 横向拼接）

### 4. 企业规则自动匹配
- 在"备注"列后自动新增40+分类匹配列
- 根据"项目特征"和"项目名称"关键词自动分类
- 匹配列包括：砌体工程、砼工程、钢筋工程、模板工程、防水工程等

## 安装

```bash
cd cost_excel_processor
pip install -r requirements.txt
```

## 使用方式

### 方式一：命令行

```bash
python -m cost_excel_processor.cli <输入Excel文件路径> [输出Excel文件路径]
```

示例：
```bash
python -m cost_excel_processor.cli ./测试清单_非标准格式.xlsx
```

### 方式二：Web界面（推荐）

```bash
streamlit run cost_excel_processor/app.py
```

### 方式三：Python代码调用

```python
from cost_excel_processor.main_processor import CostExcelProcessor

processor = CostExcelProcessor()
processor.load("input.xlsx")
processor.analyze()
result_df = processor.process()
processor.save("output.xlsx")
```

## 处理流程

```
上传Excel
    ↓
[Step 1] 读取所有工作表，识别结构（单表/多表，合并单元格检测）
    ↓
[Step 2] 解析表头：处理多级合并表头 → 标准化一级表头
    ↓
[Step 3] 多表整合：字段对齐、缺失值填充、纵向合并、去重
    ↓
[Step 4] 企业规则匹配：40+分类列关键词自动匹配
    ↓
[Step 5] 数据校验与清理
    ↓
输出"成本分析清单表1"（标准化Excel）
```

## 输出标准

输出Excel包含以下列（顺序固定）：

| 列名 | 说明 |
|------|------|
| 业态 | 自动从工作表名提取（如"教学楼地上"） |
| 项目名称 | 标准化 |
| 项目特征 | 标准化 |
| 计量单位 | 标准化 |
| 工程量 | 数值格式 |
| 综合单价 | 数值格式 |
| 综合合价 | 数值格式 |
| 备注 | 始终在最后一列 |
| 砌体工程 | 规则匹配列（自动新增） |
| 砂浆 | 规则匹配列（自动新增） |
| ... | 共40+规则匹配列 |

## 测试

生成测试数据并运行：

```bash
python cost_excel_processor/create_test_data.py
python -m cost_excel_processor.cli 测试清单_非标准格式.xlsx
```

## 项目结构

```
cost_excel_processor/
├── __init__.py          # 包初始化
├── __main__.py          # python -m 入口
├── config.py            # 常量配置（必填表头、规则关键词）
├── utils.py             # 工具函数
├── excel_reader.py      # Excel读取与结构识别
├── header_parser.py     # 多级表头解析
├── data_integrator.py  # 多表整合
├── rule_matcher.py      # 企业规则匹配
├── main_processor.py    # 主处理流程
├── cli.py              # 命令行接口
├── app.py              # Streamlit Web界面
├── create_test_data.py # 测试数据生成器
└── requirements.txt    # 依赖清单
```

## 版本

- v0.1.0（初稿）- 核心功能实现
