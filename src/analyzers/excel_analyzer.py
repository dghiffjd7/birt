"""
Excel文件分析器
负责解析Excel文件结构，提取报表元数据
"""
import pandas as pd
import openpyxl
from typing import Dict, List, Any, Optional
from pathlib import Path
import logging
from dataclasses import dataclass, asdict

logger = logging.getLogger(__name__)


@dataclass
class TableInfo:
    """表格信息数据类"""
    name: str
    headers: List[str]
    data_types: Dict[str, str]
    sample_data: Dict[str, Any]
    row_count: int
    has_totals: bool = False
    
    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class ChartInfo:
    """图表信息数据类"""
    chart_type: str
    title: str
    data_range: str
    series_count: int
    
    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class ExcelAnalysisResult:
    """Excel分析结果"""
    file_name: str
    tables: List[TableInfo]
    charts: List[ChartInfo]
    suggested_report_type: str
    complexity_score: int  # 1-5分，用于选择模板
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'file_name': self.file_name,
            'tables': [table.to_dict() for table in self.tables],
            'charts': [chart.to_dict() for chart in self.charts],
            'suggested_report_type': self.suggested_report_type,
            'complexity_score': self.complexity_score
        }


class ExcelAnalyzer:
    """Excel文件分析器"""
    
    def __init__(self):
        self.supported_extensions = ['.xlsx', '.xls']
    
    def analyze_excel(self, excel_path: str) -> ExcelAnalysisResult:
        """
        分析Excel文件并提取报表元数据
        
        Args:
            excel_path: Excel文件路径
            
        Returns:
            ExcelAnalysisResult: 分析结果
        """
        file_path = Path(excel_path)
        
        if not file_path.exists():
            raise FileNotFoundError(f"Excel文件不存在: {excel_path}")
        
        if file_path.suffix not in self.supported_extensions:
            raise ValueError(f"不支持的文件格式: {file_path.suffix}")
        
        logger.info(f"开始分析Excel文件: {excel_path}")
        
        try:
            # 1. 分析表格数据
            tables = self._analyze_tables(excel_path)
            
            # 2. 分析图表
            charts = self._analyze_charts(excel_path)
            
            # 3. 推断报表类型和复杂度
            report_type = self._suggest_report_type(tables, charts)
            complexity = self._calculate_complexity(tables, charts)
            
            result = ExcelAnalysisResult(
                file_name=file_path.name,
                tables=tables,
                charts=charts,
                suggested_report_type=report_type,
                complexity_score=complexity
            )
            
            logger.info(f"分析完成，发现 {len(tables)} 个表格，{len(charts)} 个图表")
            return result
            
        except Exception as e:
            logger.error(f"分析Excel文件失败: {str(e)}")
            raise
    
    def _analyze_tables(self, excel_path: str) -> List[TableInfo]:
        """分析Excel中的表格数据"""
        tables = []
        
        try:
            # 读取所有工作表
            excel_data = pd.read_excel(excel_path, sheet_name=None)
            
            for sheet_name, df in excel_data.items():
                if df.empty:
                    continue
                
                # 清理数据：删除完全为空的行和列
                df = df.dropna(how='all').dropna(axis=1, how='all')
                
                if df.empty:
                    continue
                
                # 检测是否有合计行
                has_totals = self._detect_totals(df)
                
                table_info = TableInfo(
                    name=sheet_name,
                    headers=list(df.columns),
                    data_types=self._get_simple_data_types(df),
                    sample_data=self._get_sample_data(df),
                    row_count=len(df),
                    has_totals=has_totals
                )
                
                tables.append(table_info)
                logger.debug(f"分析表格: {sheet_name}, 行数: {len(df)}, 列数: {len(df.columns)}")
                
        except Exception as e:
            logger.error(f"分析表格数据失败: {str(e)}")
            
        return tables
    
    def _analyze_charts(self, excel_path: str) -> List[ChartInfo]:
        """分析Excel中的图表"""
        charts = []
        
        try:
            workbook = openpyxl.load_workbook(excel_path)
            
            for sheet in workbook.worksheets:
                for chart in sheet._charts:
                    chart_info = ChartInfo(
                        chart_type=chart.__class__.__name__.replace('Chart', ''),
                        title=self._extract_chart_title(chart),
                        data_range=self._extract_data_range(chart),
                        series_count=len(chart.series) if hasattr(chart, 'series') else 0
                    )
                    charts.append(chart_info)
                    logger.debug(f"发现图表: {chart_info.chart_type}")
                    
        except Exception as e:
            logger.error(f"分析图表失败: {str(e)}")
            
        return charts
    
    def _detect_totals(self, df: pd.DataFrame) -> bool:
        """检测是否包含合计行"""
        total_keywords = ['合计', '总计', 'Total', 'Sum', '小计']
        
        try:
            # 将DataFrame转换为字符串进行搜索
            df_str = df.astype(str)
            
            # 检查最后几行是否包含合计关键词
            for keyword in total_keywords:
                # 检查每一列是否包含关键词
                for col in df_str.columns:
                    if df_str[col].str.contains(keyword, case=False, na=False).any():
                        return True
            
            return False
        except Exception as e:
            logger.error(f"检测合计行失败: {str(e)}")
            return False
    
    def _get_simple_data_types(self, df: pd.DataFrame) -> Dict[str, str]:
        """获取简化的数据类型"""
        type_mapping = {
            'object': 'text',
            'int64': 'number',
            'int32': 'number',
            'float64': 'decimal',
            'float32': 'decimal',
            'datetime64[ns]': 'date',
            'bool': 'boolean'
        }
        
        return {
            col: type_mapping.get(str(dtype), 'text')
            for col, dtype in df.dtypes.items()
        }
    
    def _get_sample_data(self, df: pd.DataFrame, max_rows: int = 3) -> Dict[str, Any]:
        """获取样本数据"""
        sample_df = df.head(max_rows)
        return {
            col: sample_df[col].dropna().tolist()
            for col in sample_df.columns
        }
    
    def _extract_chart_title(self, chart) -> str:
        """提取图表标题"""
        try:
            if hasattr(chart, 'title') and chart.title:
                if hasattr(chart.title, 'tx') and chart.title.tx:
                    return str(chart.title.tx)
                elif hasattr(chart.title, 'text'):
                    return str(chart.title.text)
            return "未命名图表"
        except:
            return "未命名图表"
    
    def _extract_data_range(self, chart) -> str:
        """提取图表数据范围"""
        try:
            if hasattr(chart, 'series') and chart.series:
                return str(chart.series[0].val) if chart.series[0].val else ""
            return ""
        except:
            return ""
    
    def _suggest_report_type(self, tables: List[TableInfo], charts: List[ChartInfo]) -> str:
        """推断报表类型"""
        # 简单的规则推断
        if not tables:
            return "empty"
        
        if len(charts) > 0:
            return "dashboard"  # 包含图表的仪表板
        elif len(tables) == 1 and tables[0].row_count < 50:
            return "simple_listing"  # 简单列表
        elif any(table.has_totals for table in tables):
            return "summary_report"  # 汇总报表
        else:
            return "detailed_listing"  # 详细列表
    
    def _calculate_complexity(self, tables: List[TableInfo], charts: List[ChartInfo]) -> int:
        """计算复杂度评分 (1-5)"""
        score = 1
        
        # 表格数量影响
        if len(tables) > 3:
            score += 2
        elif len(tables) > 1:
            score += 1
        
        # 图表影响
        if len(charts) > 2:
            score += 2
        elif len(charts) > 0:
            score += 1
        
        # 数据量影响
        max_rows = max((table.row_count for table in tables), default=0)
        if max_rows > 1000:
            score += 1
        
        return min(score, 5)


def test_analyzer():
    """测试函数"""
    import os
    
    # 检查是否有测试文件
    test_files = [
        "test_data.xlsx",
        "../input/sample.xlsx"
    ]
    
    analyzer = ExcelAnalyzer()
    
    for test_file in test_files:
        if os.path.exists(test_file):
            try:
                result = analyzer.analyze_excel(test_file)
                print(f"分析结果: {result.to_dict()}")
                break
            except Exception as e:
                print(f"测试失败: {e}")
    else:
        print("没有找到测试文件，请创建一个Excel文件进行测试")


if __name__ == "__main__":
    test_analyzer()