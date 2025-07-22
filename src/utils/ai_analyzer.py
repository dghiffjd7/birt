"""
AI分析器
使用OpenAI API分析Excel结构并生成BIRT配置
"""
import openai
import json
from typing import Dict, List, Any, Optional
import logging
from dataclasses import dataclass
import os
from dotenv import load_dotenv

from ..analyzers.excel_analyzer import ExcelAnalysisResult

# 加载环境变量
load_dotenv()

logger = logging.getLogger(__name__)


@dataclass
class AIAnalysisResult:
    """AI分析结果"""
    sql_query: str
    parameters: List[Dict[str, Any]]
    layout_config: Dict[str, Any]
    data_bindings: List[Dict[str, str]]
    scripts: List[Dict[str, str]]
    confidence_score: float
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'sql_query': self.sql_query,
            'parameters': self.parameters,
            'layout_config': self.layout_config,
            'data_bindings': self.data_bindings,
            'scripts': self.scripts,
            'confidence_score': self.confidence_score
        }


class AIAnalyzer:
    """AI分析器，使用OpenAI分析Excel并生成BIRT配置"""
    
    def __init__(self, 
                 api_key: Optional[str] = None,
                 base_url: Optional[str] = None,
                 model: str = "gpt-4"):
        
        self.api_key = api_key or os.getenv("OPENAI_API_KEY")
        self.base_url = base_url or os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
        self.model = model
        
        if not self.api_key:
            raise ValueError("OpenAI API密钥未设置，请在.env文件中设置OPENAI_API_KEY")
        
        # 初始化OpenAI客户端
        openai.api_key = self.api_key
        if self.base_url != "https://api.openai.com/v1":
            openai.api_base = self.base_url
        
        logger.info(f"AI分析器初始化完成，使用模型: {self.model}")
    
    async def analyze_for_birt(self, excel_result: ExcelAnalysisResult) -> AIAnalysisResult:
        """
        分析Excel结果并生成BIRT配置
        
        Args:
            excel_result: Excel分析结果
            
        Returns:
            AIAnalysisResult: AI分析结果
        """
        logger.info(f"开始AI分析: {excel_result.file_name}")
        
        try:
            # 1. 构造提示词
            prompt = self._build_prompt(excel_result)
            
            # 2. 调用OpenAI API
            response = await self._call_openai(prompt)
            
            # 3. 解析响应
            ai_result = self._parse_response(response)
            
            logger.info(f"AI分析完成，置信度: {ai_result.confidence_score}")
            return ai_result
            
        except Exception as e:
            logger.error(f"AI分析失败: {str(e)}")
            # 返回默认配置
            return self._get_fallback_result(excel_result)
    
    def _build_prompt(self, excel_result: ExcelAnalysisResult) -> str:
        """构建AI提示词"""
        
        tables_info = []
        for table in excel_result.tables:
            table_desc = f"""
表名: {table.name}
列数: {len(table.headers)}
行数: {table.row_count}
列信息: {', '.join([f"{h}({t})" for h, t in zip(table.headers, table.data_types.values())])}
样本数据: {table.sample_data}
包含合计: {table.has_totals}
"""
            tables_info.append(table_desc)
        
        charts_info = []
        for chart in excel_result.charts:
            chart_desc = f"图表类型: {chart.chart_type}, 标题: {chart.title}"
            charts_info.append(chart_desc)
        
        prompt = f"""
你是一个专业的BIRT报表开发专家，需要分析Excel文件结构并生成BIRT报表配置。

## Excel文件信息
文件名: {excel_result.file_name}
建议报表类型: {excel_result.suggested_report_type}
复杂度评分: {excel_result.complexity_score}/5

## 表格信息
{chr(10).join(tables_info)}

## 图表信息
{chr(10).join(charts_info) if charts_info else "无图表"}

## 系统信息
- 数据库: Quicksilver系统，使用JeedSQL驱动
- 主要数据表: TsUnit (单元表), TsUser (用户表), TsRole (角色表)
- 需要支持中文编码转换
- 需要添加Quicksilver特有的初始化脚本

## 生成要求
请根据上述信息生成一个完整的BIRT报表配置，包括：

1. **SQL查询语句**: 
   - 基于Quicksilver系统的表结构
   - 使用合适的JOIN和WHERE条件
   - 支持参数化查询
   - 考虑中文字段的处理

2. **报表参数**: 
   - 根据Excel内容推断可能需要的筛选参数
   - 包含参数名、类型、提示文本

3. **布局配置**: 
   - 适合的表格布局
   - 列宽和样式设置
   - 分组和汇总设置

4. **数据绑定**: 
   - 字段映射关系
   - 数据类型转换

5. **脚本代码**: 
   - Quicksilver初始化脚本
   - 数据处理脚本
   - 中文编码处理

请以JSON格式输出，结构如下：
```json
{{
    "sql_query": "SELECT ... FROM ...",
    "parameters": [
        {{
            "name": "参数名",
            "data_type": "string|integer|date",
            "prompt": "提示文本",
            "default_value": "默认值(可选)"
        }}
    ],
    "layout_config": {{
        "table_style": "simple|grouped|summary",
        "columns": [
            {{
                "name": "列名",
                "width": "宽度",
                "alignment": "left|center|right",
                "format": "格式化规则(可选)"
            }}
        ],
        "grouping": {{
            "enabled": true/false,
            "group_by": "分组字段",
            "show_totals": true/false
        }}
    }},
    "data_bindings": [
        {{
            "excel_column": "Excel列名",
            "db_field": "数据库字段名",
            "data_type": "数据类型"
        }}
    ],
    "scripts": [
        {{
            "event": "initialize|beforeOpen|onCreate",
            "code": "JavaScript代码"
        }}
    ],
    "confidence_score": 0.85
}}
```

注意事项：
- SQL查询要符合JeedSQL语法
- 参数名使用英文，提示文本使用中文
- 考虑Quicksilver系统的特殊需求
- 置信度评分范围0-1，根据Excel信息完整度评估
"""
        
        return prompt
    
    async def _call_openai(self, prompt: str) -> str:
        """调用OpenAI API"""
        try:
            response = openai.ChatCompletion.create(
                model=self.model,
                messages=[
                    {
                        "role": "system", 
                        "content": "你是一个专业的BIRT报表开发专家，擅长分析Excel文件并生成高质量的BIRT报表配置。"
                    },
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,
                max_tokens=2000
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            logger.error(f"OpenAI API调用失败: {str(e)}")
            raise
    
    def _parse_response(self, response: str) -> AIAnalysisResult:
        """解析OpenAI响应"""
        try:
            # 提取JSON部分
            json_start = response.find('{')
            json_end = response.rfind('}') + 1
            
            if json_start == -1 or json_end == 0:
                raise ValueError("响应中未找到有效的JSON")
            
            json_str = response[json_start:json_end]
            data = json.loads(json_str)
            
            return AIAnalysisResult(
                sql_query=data.get('sql_query', 'SELECT 1 as id'),
                parameters=data.get('parameters', []),
                layout_config=data.get('layout_config', {}),
                data_bindings=data.get('data_bindings', []),
                scripts=data.get('scripts', []),
                confidence_score=data.get('confidence_score', 0.5)
            )
            
        except json.JSONDecodeError as e:
            logger.error(f"JSON解析失败: {str(e)}")
            logger.debug(f"原始响应: {response}")
            raise ValueError(f"AI响应格式错误: {str(e)}")
        except Exception as e:
            logger.error(f"响应解析失败: {str(e)}")
            raise
    
    def _get_fallback_result(self, excel_result: ExcelAnalysisResult) -> AIAnalysisResult:
        """获取默认/回退配置"""
        logger.warning("使用回退配置")
        
        # 基于Excel分析结果生成简单的默认配置
        if excel_result.tables:
            main_table = excel_result.tables[0]
            
            # 生成简单的SQL
            columns = ", ".join([f"'{header}' as {self._clean_name(header)}" 
                               for header in main_table.headers])
            sql_query = f"SELECT {columns}"
            
            # 生成列配置
            columns_config = []
            for header in main_table.headers:
                columns_config.append({
                    "name": header,
                    "width": "100px",
                    "alignment": "left"
                })
        else:
            sql_query = "SELECT 1 as id, '示例数据' as name"
            columns_config = [
                {"name": "ID", "width": "50px", "alignment": "center"},
                {"name": "名称", "width": "150px", "alignment": "left"}
            ]
        
        return AIAnalysisResult(
            sql_query=sql_query,
            parameters=[],
            layout_config={
                "table_style": "simple",
                "columns": columns_config,
                "grouping": {"enabled": False}
            },
            data_bindings=[],
            scripts=[
                {
                    "event": "initialize",
                    "code": """
importPackage(Packages.com.jeedsoft.quicksilver.report.util);
ReportUtil.initializeDataSource(reportContext, this, true);
"""
                }
            ],
            confidence_score=0.3
        )
    
    def _clean_name(self, name: str) -> str:
        """清理名称，移除特殊字符"""
        import re
        return re.sub(r'[^\w]', '_', name)


# 同步版本的简化接口
class SimpleAIAnalyzer:
    """简化的AI分析器，提供同步接口"""
    
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.getenv("OPENAI_API_KEY")
        if not self.api_key:
            logger.warning("未设置OpenAI API密钥，将使用模拟分析")
    
    def analyze_excel_simple(self, excel_result: ExcelAnalysisResult) -> Dict[str, Any]:
        """简单的Excel分析，返回基础配置"""
        
        if not self.api_key:
            return self._get_mock_result(excel_result)
        
        # 这里可以实现简单的规则基础分析
        # 暂时返回模拟结果
        return self._get_mock_result(excel_result)
    
    def _get_mock_result(self, excel_result: ExcelAnalysisResult) -> Dict[str, Any]:
        """获取模拟分析结果"""
        
        if excel_result.tables:
            main_table = excel_result.tables[0]
            
            # 生成基于表结构的SQL
            if "编码" in main_table.headers and "名称" in main_table.headers:
                sql_query = "SELECT FCode as 编码, FName as 名称 FROM TsUnit ORDER BY FCode"
            else:
                sql_query = f"SELECT {', '.join(main_table.headers[:5])} FROM YourTable"
            
            parameters = []
            if main_table.row_count > 100:  # 大数据量时添加筛选参数
                parameters.append({
                    "name": "filter_param",
                    "data_type": "string",
                    "prompt": "请输入筛选条件"
                })
        else:
            sql_query = "SELECT FCode as 编码, FName as 名称 FROM TsUnit ORDER BY FCode"
            parameters = []
        
        return {
            'sql_query': sql_query,
            'parameters': parameters,
            'report_title': excel_result.file_name.replace('.xlsx', '').replace('.xls', ''),
            'db_url': 'jdbc:jeedsql:jtds:sqlserver://localhost:1433;DatabaseName=QS;',
            'db_user': '',
            'db_password': ''
        }


def test_ai_analyzer():
    """测试AI分析器"""
    from ..analyzers.excel_analyzer import ExcelAnalyzer, TableInfo, ExcelAnalysisResult
    
    # 创建测试数据
    test_table = TableInfo(
        name="TestSheet",
        headers=["编码", "名称", "类型", "数量"],
        data_types={"编码": "text", "名称": "text", "类型": "text", "数量": "number"},
        sample_data={
            "编码": ["001", "002", "003"], 
            "名称": ["项目1", "项目2", "项目3"],
            "类型": ["A", "B", "C"],
            "数量": [10, 20, 30]
        },
        row_count=100
    )
    
    test_result = ExcelAnalysisResult(
        file_name="测试报表.xlsx",
        tables=[test_table],
        charts=[],
        suggested_report_type="simple_listing",
        complexity_score=3
    )
    
    # 测试简单分析器
    analyzer = SimpleAIAnalyzer()
    result = analyzer.analyze_excel_simple(test_result)
    print("简单分析结果:")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    test_ai_analyzer()