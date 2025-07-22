"""
BIRT报表生成器
根据Excel分析结果生成.rptdesign文件
"""
import xml.etree.ElementTree as ET
from typing import Dict, List, Any, Optional
from pathlib import Path
import logging
from jinja2 import Template
import uuid
from datetime import datetime
import xml.dom.minidom

import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent.parent))

from analyzers.excel_analyzer import ExcelAnalysisResult, TableInfo

logger = logging.getLogger(__name__)


class BIRTGenerator:
    """BIRT报表生成器"""
    
    def __init__(self, template_dir: str = "templates/birt"):
        self.template_dir = Path(template_dir)
        self.namespace = {"": "http://www.eclipse.org/birt/2005/design"}
        
    def generate_report(self, 
                       analysis_result: ExcelAnalysisResult,
                       ai_config: Dict[str, Any],
                       output_path: str) -> str:
        """
        生成BIRT报表文件
        
        Args:
            analysis_result: Excel分析结果
            ai_config: AI生成的配置
            output_path: 输出文件路径
            
        Returns:
            str: 生成的报表文件路径
        """
        logger.info(f"开始生成BIRT报表: {analysis_result.file_name}")
        
        try:
            # 1. 选择合适的模板
            template_name = self._select_template(analysis_result)
            
            # 2. 加载基础模板
            template_content = self._load_template(template_name)
            
            # 3. 生成报表XML
            report_xml = self._generate_report_xml(
                template_content, 
                analysis_result, 
                ai_config
            )
            
            # 4. 美化XML格式
            formatted_xml = self._format_xml(report_xml)
            
            # 5. 保存文件
            output_file = self._save_report(formatted_xml, output_path, analysis_result.file_name)
            
            logger.info(f"报表生成成功: {output_file}")
            return output_file
            
        except Exception as e:
            logger.error(f"生成BIRT报表失败: {str(e)}")
            raise
    
    def _select_template(self, analysis_result: ExcelAnalysisResult) -> str:
        """根据分析结果选择模板"""
        template_map = {
            "simple_listing": "simple_listing_template.xml",
            "detailed_listing": "detailed_listing_template.xml", 
            "summary_report": "summary_report_template.xml",
            "dashboard": "dashboard_template.xml",
            "empty": "base_template.xml"
        }
        
        template_name = template_map.get(
            analysis_result.suggested_report_type, 
            "base_template.xml"
        )
        
        logger.debug(f"选择模板: {template_name}")
        return template_name
    
    def _load_template(self, template_name: str) -> str:
        """加载模板文件"""
        template_path = self.template_dir / template_name
        
        # 如果特定模板不存在，使用基础模板
        if not template_path.exists():
            template_path = self.template_dir / "base_template.xml"
            logger.warning(f"模板文件不存在，使用基础模板: {template_path}")
        
        with open(template_path, 'r', encoding='utf-8') as f:
            return f.read()
    
    def _generate_report_xml(self, 
                           template_content: str,
                           analysis_result: ExcelAnalysisResult,
                           ai_config: Dict[str, Any]) -> str:
        """生成完整的报表XML"""
        
        # 1. 准备模板变量
        template_vars = self._prepare_template_vars(analysis_result, ai_config)
        
        # 2. 渲染基础模板
        template = Template(template_content)
        base_xml = template.render(**template_vars)
        
        # 3. 解析XML并添加详细内容
        root = ET.fromstring(base_xml)
        
        # 4. 添加自定义JavaScript函数
        self._add_custom_functions(root)
        
        # 5. 添加数据集
        self._add_datasets(root, analysis_result, ai_config)
        
        # 6. 添加专业报表布局
        self._add_professional_report_layout(root, analysis_result, ai_config)
        
        # 7. 添加参数（如果有）
        if ai_config.get('parameters'):
            self._add_parameters(root, ai_config['parameters'])
        
        # 生成XML字符串
        xml_str = ET.tostring(root, encoding='unicode')
        
        # 添加XML声明
        if not xml_str.startswith('<?xml'):
            xml_str = '<?xml version="1.0" encoding="UTF-8"?>\n' + xml_str
        
        # 移除ns0前缀，使用标准BIRT格式
        xml_str = xml_str.replace('ns0:', '')
        xml_str = xml_str.replace(' xmlns:ns0=', ' xmlns=')
        
        # 修复编码问题 - 确保HTML和JavaScript内容正确
        xml_str = self._fix_encoding_issues(xml_str)
        
        return xml_str
    
    def _prepare_template_vars(self, 
                             analysis_result: ExcelAnalysisResult,
                             ai_config: Dict[str, Any]) -> Dict[str, Any]:
        """准备模板变量"""
        return {
            'DB_URL': ai_config.get('db_url', 'jdbc:jeedsql:jtds:sqlserver://localhost:1433;DatabaseName=QS;'),
            'DB_USER': ai_config.get('db_user', ''),
            'DB_PASSWORD': ai_config.get('db_password', ''),
            'REPORT_TITLE': ai_config.get('report_title', analysis_result.file_name.replace('.xlsx', '')),
            'CREATED_DATE': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'GENERATOR_VERSION': 'AI BIRT Generator v1.0'
        }
    
    def _format_xml(self, xml_content: str) -> str:
        """美化XML格式，提高可读性"""
        try:
            # 解析XML
            dom = xml.dom.minidom.parseString(xml_content)
            
            # 美化格式
            formatted = dom.toprettyxml(indent="    ", encoding="UTF-8")
            
            # 转换为字符串并清理空行
            if isinstance(formatted, bytes):
                formatted = formatted.decode('utf-8')
            
            # 移除多余的空行
            lines = [line for line in formatted.split('\n') if line.strip()]
            
            return '\n'.join(lines)
            
        except Exception as e:
            logger.warning(f"XML格式化失败，使用原始格式: {e}")
            return xml_content

    def _add_datasets(self, 
                     root: ET.Element,
                     analysis_result: ExcelAnalysisResult,
                     ai_config: Dict[str, Any]):
        """添加数据集 - 改进版本"""
        # 使用命名空间查找
        ns = {'': 'http://www.eclipse.org/birt/2005/design'}
        data_sets = root.find('.//data-sets', ns)
        if data_sets is None:
            data_sets = ET.SubElement(root, 'data-sets')
        
        # 主数据集
        main_dataset = ET.SubElement(data_sets, 'oda-data-set')
        main_dataset.set('extensionID', 'org.eclipse.birt.report.data.oda.jdbc.JdbcSelectDataSet')
        main_dataset.set('name', 'MainDataSet')
        main_dataset.set('id', str(self._generate_id()))
        
        # SQL查询 - 使用双引号包围字段名
        sql_query = ai_config.get('sql_query', 'SELECT 1 as id, \'示例数据\' as name')
        # 确保字段名使用双引号
        sql_query = self._normalize_sql_quotes(sql_query)
        
        query_text = ET.SubElement(main_dataset, 'property')
        query_text.set('name', 'queryText')
        query_text.text = sql_query
        
        # 数据源引用
        data_source = ET.SubElement(main_dataset, 'property')
        data_source.set('name', 'dataSource')
        data_source.text = 'default'
        
        # 添加参数绑定（如果有参数）
        if ai_config.get('parameters'):
            self._add_parameter_bindings(main_dataset, ai_config['parameters'])
        
        # 添加列结构
        self._add_result_set_structure(main_dataset, analysis_result, ai_config)
        
        logger.debug("已添加主数据集")

    def _normalize_sql_quotes(self, sql: str) -> str:
        """规范化SQL中的引号，确保字段名使用双引号"""
        import re
        
        # 将单引号包围的字段名替换为双引号
        # 匹配模式：'字段名' -> "字段名"
        pattern = r"'([^']*?)'"
        
        def replace_quotes(match):
            content = match.group(1)
            # 如果是中文字段名或包含特殊字符，使用双引号
            if any('\u4e00' <= char <= '\u9fff' for char in content) or ' ' in content:
                return f'"{content}"'
            return match.group(0)
        
        return re.sub(pattern, replace_quotes, sql)

    def _add_parameter_bindings(self, dataset: ET.Element, parameters: List[Dict[str, Any]]):
        """添加参数绑定"""
        bindings = ET.SubElement(dataset, 'list-property')
        bindings.set('name', 'parameterBindings')
        
        for i, param in enumerate(parameters):
            structure = ET.SubElement(bindings, 'structure')
            
            param_name = ET.SubElement(structure, 'property')
            param_name.set('name', 'paramName')
            param_name.text = param.get('name', f'param{i+1}')
            
            position = ET.SubElement(structure, 'property')
            position.set('name', 'position')
            position.text = str(i + 1)

    def _add_result_set_structure(self,
                                 dataset: ET.Element,
                                 analysis_result: ExcelAnalysisResult,
                                 ai_config: Dict[str, Any]):
        """添加完整的结果集结构"""
        # 添加columnHints
        if analysis_result.tables:
            self._add_column_hints(dataset, analysis_result.tables[0].headers)
        
        # 添加cachedMetaData
        self._add_cached_metadata(dataset, analysis_result)
        
        # 添加标准resultSet
        result_set_list = ET.SubElement(dataset, 'list-property')
        result_set_list.set('name', 'resultSet')
        
        # 如果有表格数据，使用第一个表格的列结构
        if analysis_result.tables:
            main_table = analysis_result.tables[0]
            
            for i, (header, data_type) in enumerate(zip(main_table.headers, main_table.data_types.values())):
                result_set_column = ET.SubElement(result_set_list, 'structure')
                
                # 位置
                position = ET.SubElement(result_set_column, 'property')
                position.set('name', 'position')
                position.text = str(i + 1)
                
                # 列名
                name_prop = ET.SubElement(result_set_column, 'property')
                name_prop.set('name', 'name')
                name_prop.text = header
                
                # 原生列名
                native_name = ET.SubElement(result_set_column, 'property')
                native_name.set('name', 'nativeName')
                native_name.text = header
                
                # 数据类型
                data_type_prop = ET.SubElement(result_set_column, 'property')
                data_type_prop.set('name', 'dataType')
                data_type_prop.text = self._map_data_type(data_type)
                
                # 原生数据类型
                native_type = ET.SubElement(result_set_column, 'property')
                native_type.set('name', 'nativeDataType')
                native_type.text = self._get_jdbc_type_code(data_type)

    def _get_jdbc_type_code(self, data_type: str) -> str:
        """获取JDBC类型代码"""
        type_codes = {
            'string': '12',  # VARCHAR
            'integer': '4',  # INTEGER
            'float': '6',    # FLOAT
            'date': '91',    # DATE
            'boolean': '16'  # BOOLEAN
        }
        birt_type = self._map_data_type(data_type)
        return type_codes.get(birt_type, '12')
    
    def _add_report_layout(self,
                          root: ET.Element,
                          analysis_result: ExcelAnalysisResult,
                          ai_config: Dict[str, Any]):
        """添加报表布局"""
        body = root.find('.//body')
        if body is None:
            body = ET.SubElement(root, 'body')
        
        # 根据报表类型选择布局
        if analysis_result.suggested_report_type == "simple_listing":
            self._add_simple_table_layout(body, analysis_result)
        elif analysis_result.suggested_report_type == "summary_report":
            self._add_summary_layout(body, analysis_result)
        else:
            self._add_default_layout(body, analysis_result)
    
    def _add_simple_table_layout(self, body: ET.Element, analysis_result: ExcelAnalysisResult):
        """添加简单表格布局"""
        if not analysis_result.tables:
            return
        
        main_table = analysis_result.tables[0]
        
        # 创建表格元素
        table = ET.SubElement(body, 'table')
        table.set('id', str(self._generate_id()))
        table.set('name', 'MainTable')
        
        # 表格属性
        table_property = ET.SubElement(table, 'property')
        table_property.set('name', 'dataSet')
        table_property.text = 'MainDataSet'
        
        # 列定义
        for header in main_table.headers:
            column = ET.SubElement(table, 'column')
            column.set('id', str(self._generate_id()))
        
        # 表头
        header_row = ET.SubElement(table, 'header')
        header_tr = ET.SubElement(header_row, 'row')
        header_tr.set('id', str(self._generate_id()))
        
        for header in main_table.headers:
            cell = ET.SubElement(header_tr, 'cell')
            cell.set('id', str(self._generate_id()))
            
            label = ET.SubElement(cell, 'label')
            label.set('id', str(self._generate_id()))
            
            text_prop = ET.SubElement(label, 'property')
            text_prop.set('name', 'text')
            text_prop.text = header
        
        # 详细行
        detail = ET.SubElement(table, 'detail')
        detail_tr = ET.SubElement(detail, 'row')
        detail_tr.set('id', str(self._generate_id()))
        
        for i, header in enumerate(main_table.headers):
            cell = ET.SubElement(detail_tr, 'cell')
            cell.set('id', str(self._generate_id()))
            
            data_item = ET.SubElement(cell, 'data')
            data_item.set('id', str(self._generate_id()))
            
            result_expr = ET.SubElement(data_item, 'property')
            result_expr.set('name', 'resultSetExpression')
            result_expr.text = f'row["{self._clean_column_name(header)}"]'
        
        logger.debug("已添加简单表格布局")
    
    def _add_summary_layout(self, body: ET.Element, analysis_result: ExcelAnalysisResult):
        """添加汇总报表布局"""
        # 暂时使用简单布局
        self._add_simple_table_layout(body, analysis_result)
    
    def _add_default_layout(self, body: ET.Element, analysis_result: ExcelAnalysisResult):
        """添加默认布局"""
        # 添加标题
        title_label = ET.SubElement(body, 'label')
        title_label.set('id', str(self._generate_id()))
        
        title_text = ET.SubElement(title_label, 'property')
        title_text.set('name', 'text')
        title_text.text = f"报表: {analysis_result.file_name}"
        
        # 如果有表格，添加表格
        if analysis_result.tables:
            self._add_simple_table_layout(body, analysis_result)
    
    def _add_parameters(self, root: ET.Element, parameters: List[Dict[str, Any]]):
        """添加报表参数 - 改进版本"""
        # 创建parameters节点（在data-sources之前）
        params_node = ET.Element('parameters')
        
        # 插入到正确位置
        data_sources = root.find('.//data-sources')
        if data_sources is not None:
            root.insert(list(root).index(data_sources), params_node)
        else:
            root.insert(0, params_node)
        
        for param in parameters:
            scalar_param = ET.SubElement(params_node, 'scalar-parameter')
            scalar_param.set('name', param.get('name', 'param1'))
            scalar_param.set('id', str(self._generate_id()))
            
            # 数据类型
            data_type = ET.SubElement(scalar_param, 'property')
            data_type.set('name', 'dataType')
            data_type.text = param.get('data_type', 'string')
            
            # 提示文本
            prompt_text = ET.SubElement(scalar_param, 'property')
            prompt_text.set('name', 'promptText')
            prompt_text.text = param.get('prompt', '请输入参数值')
            
            # 是否必需
            is_required = ET.SubElement(scalar_param, 'property')
            is_required.set('name', 'isRequired')
            is_required.text = str(param.get('required', False)).lower()
            
            # 控件类型
            control_type = ET.SubElement(scalar_param, 'property')
            control_type.set('name', 'controlType')
            control_type.text = param.get('control_type', 'text-box')
    
    def _clean_column_name(self, name: str) -> str:
        """清理列名，确保符合BIRT要求"""
        # 移除特殊字符，替换为下划线
        import re
        clean_name = re.sub(r'[^\w\u4e00-\u9fff]', '_', name)
        return clean_name
    
    def _map_data_type(self, data_type: str) -> str:
        """映射数据类型到BIRT类型"""
        type_map = {
            'text': 'string',
            'number': 'integer',
            'decimal': 'float',
            'date': 'date',
            'boolean': 'boolean'
        }
        return type_map.get(data_type, 'string')
    
    def _generate_id(self) -> int:
        """生成唯一ID"""
        return int(str(uuid.uuid4().int)[:8])
    
    def _save_report(self, xml_content: str, output_path: str, file_name: str) -> str:
        """保存报表文件"""
        output_dir = Path(output_path)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # 生成输出文件名
        base_name = Path(file_name).stem
        output_file = output_dir / f"{base_name}.rptdesign"
        
        # 保存文件
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(xml_content)
        
        return str(output_file)

    def _add_custom_functions(self, root: ET.Element):
        """添加自定义JavaScript函数"""
        # 在根元素下添加初始化方法
        existing_method = root.find('.//method[@name="initialize"]')
        if existing_method is not None:
            root.remove(existing_method)
        
        method = ET.Element('method')
        method.set('name', 'initialize')
        
        # JavaScript代码
        js_code = '''importPackage(Packages.java.text);
 
function day() {
    var fmt = new SimpleDateFormat("yyyy年MM月dd日");
    return fmt.format(new Date());
}

function formatDate(date, pattern) {
    if (!date) return "";
    var fmt = new SimpleDateFormat(pattern || "yyyy-MM-dd");
    return fmt.format(date);
}

function formatNumber(number, pattern) {
    if (!number) return "0";
    var fmt = new DecimalFormat(pattern || "#,##0.00");
    return fmt.format(number);
}'''
        
        method.text = js_code
        
        # 插入到合适位置（在property元素之后）
        properties = root.findall('./property')
        if properties:
            insert_index = list(root).index(properties[-1]) + 1
            root.insert(insert_index, method)
        else:
            root.insert(0, method)

    def _add_professional_report_layout(self, root: ET.Element, analysis_result: ExcelAnalysisResult, ai_config: Dict[str, Any]):
        """添加专业的报表布局"""
        body = root.find('.//body')
        if body is None:
            body = ET.SubElement(root, 'body')
        
        # 清空现有内容
        body.clear()
        
        # 1. 添加专业标题
        report_title = ai_config.get('report_title', f"数据报表: {analysis_result.file_name}")
        self._add_professional_header(body, report_title)
        
        # 2. 添加日期信息区域
        self._add_date_info_section(body)
        
        # 3. 添加带样式的数据表格
        if analysis_result.tables:
            self._add_styled_table_layout(body, analysis_result)
    
    def _add_professional_header(self, body: ET.Element, report_title: str):
        """添加专业的报表标题区域"""
        title_grid = ET.SubElement(body, 'grid')
        title_grid.set('id', str(self._generate_id()))
        
        # 设置边框样式和宽度
        self._set_border_style(title_grid)
        title_grid_width = ET.SubElement(title_grid, 'property')
        title_grid_width.set('name', 'width')
        title_grid_width.text = '600pt'
        
        # 添加列定义
        for i in range(5):
            column = ET.SubElement(title_grid, 'column')
            column.set('id', str(self._generate_id()))
        
        # 标题行
        row = ET.SubElement(title_grid, 'row')
        row.set('id', str(self._generate_id()))
        
        cell = ET.SubElement(row, 'cell')
        cell.set('id', str(self._generate_id()))
        cell_colspan = ET.SubElement(cell, 'property')
        cell_colspan.set('name', 'colSpan')
        cell_colspan.text = '5'
        
        text_elem = ET.SubElement(cell, 'text')
        text_elem.set('id', str(self._generate_id()))
        
        # 字体设置
        font_family = ET.SubElement(text_elem, 'property')
        font_family.set('name', 'fontFamily')
        font_family.text = '"標楷體"'
        
        margin_top = ET.SubElement(text_elem, 'property')
        margin_top.set('name', 'marginTop')
        margin_top.text = '0pt'
        
        text_align = ET.SubElement(text_elem, 'property')
        text_align.set('name', 'textAlign')
        text_align.text = 'center'
        
        content_type = ET.SubElement(text_elem, 'property')
        content_type.set('name', 'contentType')
        content_type.text = 'html'
        
        content = ET.SubElement(text_elem, 'text-property')
        content.set('name', 'content')
        # 使用CDATA来包装HTML内容，避免编码问题
        html_content = f'<H1><B><U>{report_title}</U></B></H1>'
        content.text = html_content

    def _add_date_info_section(self, body: ET.Element):
        """添加日期信息区域"""
        date_grid = ET.SubElement(body, 'grid')
        date_grid.set('id', str(self._generate_id()))
        
        # 设置边框
        self._set_border_style(date_grid)
        
        # 添加5列，每列120pt
        for i in range(5):
            column = ET.SubElement(date_grid, 'column')
            column.set('id', str(self._generate_id()))
            width_prop = ET.SubElement(column, 'property')
            width_prop.set('name', 'width')
            width_prop.text = '120pt'
        
        # 日期行
        row = ET.SubElement(date_grid, 'row')
        row.set('id', str(self._generate_id()))
        
        # 资料日期
        data_date_cell = ET.SubElement(row, 'cell')
        data_date_cell.set('id', str(self._generate_id()))
        
        data_date_text = ET.SubElement(data_date_cell, 'text-data')
        data_date_text.set('id', str(self._generate_id()))
        
        font_family = ET.SubElement(data_date_text, 'property')
        font_family.set('name', 'fontFamily')
        font_family.text = '"標楷體"'
        
        value_expr = ET.SubElement(data_date_text, 'expression')
        value_expr.set('name', 'valueExpr')
        # 确保JavaScript表达式不被编码
        value_expr.text = '"資料日期：" + day() + " ～ " + day()'
        
        content_type = ET.SubElement(data_date_text, 'property')
        content_type.set('name', 'contentType')
        content_type.text = 'html'
        
        # 空白单元格
        for i in range(3):
            empty_cell = ET.SubElement(row, 'cell')
            empty_cell.set('id', str(self._generate_id()))
        
        # 制表日期
        create_date_cell = ET.SubElement(row, 'cell')
        create_date_cell.set('id', str(self._generate_id()))
        
        create_date_text = ET.SubElement(create_date_cell, 'text-data')
        create_date_text.set('id', str(self._generate_id()))
        
        font_family2 = ET.SubElement(create_date_text, 'property')
        font_family2.set('name', 'fontFamily')
        font_family2.text = '"標楷體"'
        
        text_align = ET.SubElement(create_date_text, 'property')
        text_align.set('name', 'textAlign')
        text_align.text = 'right'
        
        value_expr2 = ET.SubElement(create_date_text, 'expression')
        value_expr2.set('name', 'valueExpr')
        # 确保JavaScript表达式不被编码
        value_expr2.text = '"製表日期："+day()'
        
        content_type2 = ET.SubElement(create_date_text, 'property')
        content_type2.set('name', 'contentType')
        content_type2.text = 'html'

    def _set_border_style(self, element: ET.Element):
        """为元素添加标准边框样式"""
        border_styles = [
            ('borderBottomStyle', 'solid'),
            ('borderBottomWidth', 'thin'),
            ('borderLeftStyle', 'solid'),
            ('borderLeftWidth', 'thin'),
            ('borderRightStyle', 'solid'),
            ('borderRightWidth', 'thin'),
            ('borderTopStyle', 'solid'),
            ('borderTopWidth', 'thin')
        ]
        
        for style_name, style_value in border_styles:
            style_prop = ET.SubElement(element, 'property')
            style_prop.set('name', style_name)
            style_prop.text = style_value

    def _add_styled_table_layout(self, body: ET.Element, analysis_result: ExcelAnalysisResult):
        """添加带样式的表格布局"""
        if not analysis_result.tables:
            return
        
        main_table = analysis_result.tables[0]
        
        # 创建表格
        table = ET.SubElement(body, 'table')
        table.set('id', str(self._generate_id()))
        table.set('name', 'MainTable')
        
        # 数据集绑定
        dataset_prop = ET.SubElement(table, 'property')
        dataset_prop.set('name', 'dataSet')
        dataset_prop.text = 'MainDataSet'
        
        # 添加boundDataColumns
        self._add_bound_data_columns(table, main_table)
        
        # 列定义（带宽度）
        for i, header in enumerate(main_table.headers):
            column = ET.SubElement(table, 'column')
            column.set('id', str(self._generate_id()))
            width_prop = ET.SubElement(column, 'property')
            width_prop.set('name', 'width')
            width_prop.text = '120pt'
        
        # 表头（带样式）
        header_section = ET.SubElement(table, 'header')
        header_row = ET.SubElement(header_section, 'row')
        header_row.set('id', str(self._generate_id()))
        
        for header in main_table.headers:
            cell = ET.SubElement(header_row, 'cell')
            cell.set('id', str(self._generate_id()))
            
            # 添加边框
            self._set_border_style(cell)
            
            label = ET.SubElement(cell, 'label')
            label.set('id', str(self._generate_id()))
            
            # 字体设置
            font_family = ET.SubElement(label, 'property')
            font_family.set('name', 'fontFamily')
            font_family.text = '"標楷體"'
            
            text_prop = ET.SubElement(label, 'text-property')
            text_prop.set('name', 'text')
            text_prop.text = header
        
        # 数据行（带样式）
        detail_section = ET.SubElement(table, 'detail')
        detail_row = ET.SubElement(detail_section, 'row')
        detail_row.set('id', str(self._generate_id()))
        
        for header in main_table.headers:
            cell = ET.SubElement(detail_row, 'cell')
            cell.set('id', str(self._generate_id()))
            
            # 添加边框
            self._set_border_style(cell)
            
            data_item = ET.SubElement(cell, 'data')
            data_item.set('id', str(self._generate_id()))
            
            result_prop = ET.SubElement(data_item, 'property')
            result_prop.set('name', 'resultSetColumn')
            result_prop.text = header
        
        # 添加空白表尾（与手工文件一致）
        footer_section = ET.SubElement(table, 'footer')
        footer_row = ET.SubElement(footer_section, 'row')
        footer_row.set('id', str(self._generate_id()))
        
        for header in main_table.headers:
            cell = ET.SubElement(footer_row, 'cell')
            cell.set('id', str(self._generate_id()))
            self._set_border_style(cell)

    def _add_bound_data_columns(self, table: ET.Element, main_table):
        """添加绑定数据列定义"""
        bound_columns = ET.SubElement(table, 'list-property')
        bound_columns.set('name', 'boundDataColumns')
        
        for header, data_type in zip(main_table.headers, main_table.data_types.values()):
            structure = ET.SubElement(bound_columns, 'structure')
            
            # 列名
            name_prop = ET.SubElement(structure, 'property')
            name_prop.set('name', 'name')
            name_prop.text = header
            
            # 显示名
            display_name = ET.SubElement(structure, 'text-property')
            display_name.set('name', 'displayName')
            display_name.text = header
            
            # 表达式
            expression = ET.SubElement(structure, 'expression')
            expression.set('name', 'expression')
            expression.set('type', 'javascript')
            expression.text = f'dataSetRow["{header}"]'
            
            # 数据类型
            data_type_prop = ET.SubElement(structure, 'property')
            data_type_prop.set('name', 'dataType')
            data_type_prop.text = self._map_data_type(data_type)

    def _add_column_hints(self, dataset: ET.Element, headers: List[str]):
        """添加列提示信息"""
        column_hints = ET.SubElement(dataset, 'list-property')
        column_hints.set('name', 'columnHints')
        
        for header in headers:
            structure = ET.SubElement(column_hints, 'structure')
            
            # 列名
            col_name = ET.SubElement(structure, 'property')
            col_name.set('name', 'columnName')
            col_name.text = header
            
            # 分析类型
            analysis = ET.SubElement(structure, 'property')
            analysis.set('name', 'analysis')
            analysis.text = 'dimension'
            
            # 显示名
            display_name = ET.SubElement(structure, 'text-property')
            display_name.set('name', 'displayName')
            display_name.text = header
            
            # 标题
            heading = ET.SubElement(structure, 'text-property')
            heading.set('name', 'heading')
            heading.text = header

    def _add_cached_metadata(self, dataset: ET.Element, analysis_result: ExcelAnalysisResult):
        """添加缓存元数据"""
        if not analysis_result.tables:
            return
        
        main_table = analysis_result.tables[0]
        
        cached_meta = ET.SubElement(dataset, 'structure')
        cached_meta.set('name', 'cachedMetaData')
        
        result_set = ET.SubElement(cached_meta, 'list-property')
        result_set.set('name', 'resultSet')
        
        for i, (header, data_type) in enumerate(zip(main_table.headers, main_table.data_types.values())):
            structure = ET.SubElement(result_set, 'structure')
            
            position = ET.SubElement(structure, 'property')
            position.set('name', 'position')
            position.text = str(i + 1)
            
            name = ET.SubElement(structure, 'property')
            name.set('name', 'name')
            name.text = header
            
            data_type_prop = ET.SubElement(structure, 'property')
            data_type_prop.set('name', 'dataType')
            data_type_prop.text = self._map_data_type(data_type)

    def _fix_encoding_issues(self, xml_content: str) -> str:
        """修复XML内容中的编码问题"""
        # 在特定元素中避免过度编码
        # 对于JavaScript表达式，确保引号不被编码
        xml_content = xml_content.replace('&quot;資料日期：&quot;', '"資料日期："')
        xml_content = xml_content.replace('&quot;製表日期：&quot;', '"製表日期："')
        xml_content = xml_content.replace('&quot; ～ &quot;', '" ～ "')
        
        # 对于HTML内容，确保标签不被编码
        xml_content = xml_content.replace('&lt;H1&gt;&lt;B&gt;&lt;U&gt;', '<H1><B><U>')
        xml_content = xml_content.replace('&lt;/U&gt;&lt;/B&gt;&lt;/H1&gt;', '</U></B></H1>')
        
        return xml_content


def test_generator():
    """测试函数"""
    from ..analyzers.excel_analyzer import ExcelAnalyzer, TableInfo
    
    # 创建测试数据
    test_table = TableInfo(
        name="TestSheet",
        headers=["编码", "名称", "类型"],
        data_types={"编码": "text", "名称": "text", "类型": "text"},
        sample_data={"编码": ["001", "002"], "名称": ["项目1", "项目2"], "类型": ["A", "B"]},
        row_count=10
    )
    
    from ..analyzers.excel_analyzer import ExcelAnalysisResult
    test_result = ExcelAnalysisResult(
        file_name="test.xlsx",
        tables=[test_table],
        charts=[],
        suggested_report_type="simple_listing",
        complexity_score=2
    )
    
    test_config = {
        'sql_query': 'SELECT FCode as 编码, FName as 名称, FTable as 类型 FROM TsUnit ORDER BY FCode',
        'db_url': 'jdbc:jeedsql:jtds:sqlserver://localhost:1433;DatabaseName=QS;'
    }
    
    generator = BIRTGenerator()
    output_file = generator.generate_report(test_result, test_config, "output")
    print(f"测试报表生成完成: {output_file}")


if __name__ == "__main__":
    test_generator()