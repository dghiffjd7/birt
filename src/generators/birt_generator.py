"""
BIRT报表生成器 - 修复版本
解决XML结构、编码转义和命名空间问题
"""
import xml.etree.ElementTree as ET
from typing import Dict, List, Any, Optional
from pathlib import Path
import logging
from jinja2 import Template
import uuid
from datetime import datetime
import html

import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent.parent))

from analyzers.excel_analyzer import ExcelAnalysisResult, TableInfo

logger = logging.getLogger(__name__)


class BIRTGenerator:
    """BIRT报表生成器 - 修复版本"""
    
    def __init__(self, template_dir: str = "templates/birt"):
        self.template_dir = Path(template_dir)
        self.namespace = {"": "http://www.eclipse.org/birt/2005/design"}
        
    def generate_report(self, 
                       analysis_result: ExcelAnalysisResult,
                       ai_config: Dict[str, Any],
                       output_path: str) -> str:
        """生成BIRT报表文件"""
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
            
            # 4. 保存文件 (不使用minidom格式化，避免隐藏问题)
            output_file = self._save_report(report_xml, output_path, analysis_result.file_name)
            
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
        """生成完整的报表XML - 修复版本"""
        
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
        # 注册命名空间以避免 ns0: 前缀
        ET.register_namespace('', self.namespace[''])
        xml_str = ET.tostring(root, encoding='unicode', method='xml')
        
        # 添加XML声明
        if not xml_str.startswith('<?xml'):
            xml_str = '<?xml version="1.0" encoding="UTF-8"?>\n' + xml_str
        
        # >> 新的、更可靠的代码 <<
        # 2. 准备好未经转义的HTML内容
        report_title = ai_config.get('report_title', analysis_result.file_name.replace('.xlsx', ''))
        raw_html_content = f'<H1><B><U>{report_title}</U></B></H1>'
        
        # 3. 用完整的、包含原始HTML的CDATA区块替换预留位置
        cdata_block = f'<![CDATA[{raw_html_content}]]>'
        xml_str = xml_str.replace('__HTML_CONTENT_PLACEHOLDER__', cdata_block)
        
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
            'GENERATOR_VERSION': 'AI BIRT Generator Fixed v1.0'
        }

    def _add_datasets(self, 
                     root: ET.Element,
                     analysis_result: ExcelAnalysisResult,
                     ai_config: Dict[str, Any]):
        """添加数据集 - 修复版本"""
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
        
        # SQL查询
        sql_query = ai_config.get('sql_query', 'SELECT 1 as id, \'示例数据\' as name')
        
        # 修复SQL语法并获取参数顺序
        fixed_sql, ordered_params = self._fix_sql_syntax(sql_query)
        
        query_text = ET.SubElement(main_dataset, 'property')
        query_text.set('name', 'queryText')
        query_text.text = fixed_sql # 使用修复后的SQL
        
        # 数据源引用
        data_source = ET.SubElement(main_dataset, 'property')
        data_source.set('name', 'dataSource')
        data_source.text = 'default'
        
        # 添加参数绑定（使用从SQL解析出的顺序）
        if ordered_params:
            self._add_parameter_bindings(main_dataset, ordered_params)
        
        # 添加列结构
        self._add_result_set_structure(main_dataset, analysis_result, ai_config)
        
        logger.debug("已添加主数据集")

    def _add_parameter_bindings(self, dataset: ET.Element, ordered_params: List[str]):
        """根据从SQL解析出的参数名列表添加绑定"""
        bindings = ET.SubElement(dataset, 'list-property')
        bindings.set('name', 'parameterBindings')
        
        for i, param_name in enumerate(ordered_params):
            structure = ET.SubElement(bindings, 'structure')
            
            # 参数名
            param_name_prop = ET.SubElement(structure, 'property')
            param_name_prop.set('name', 'paramName')
            param_name_prop.text = param_name
            
            # 位置
            position_prop = ET.SubElement(structure, 'property')
            position_prop.set('name', 'position')
            position_prop.text = str(i + 1)

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
    
    def _fix_sql_syntax(self, sql_query: str) -> (str, List[str]):
        """
        修复SQL语法，并将命名参数转为'?'，同时返回解析出的参数名列表。
        """
        import re
        
        # 0. 验证并修复基础SQL结构
        if "SELECT 1 as id" in sql_query: # 检查是否是默认的无效查询
             return sql_query, []

        # 1. 更可靠地将所有用单引号包裹的标识符（表名.字段名）替换为双引号
        # 例如 'in'.'案號' -> "in"."案號"
        sql_query = re.sub(r"'([^']*)'\.'([^']*)'", r'"\1"."\2"', sql_query)
        sql_query = re.sub(r"FROM\s+'([^']+)'", r'FROM "\1"', sql_query, flags=re.IGNORECASE)
        sql_query = re.sub(r"JOIN\s+'([^']+)'", r'JOIN "\1"', sql_query, flags=re.IGNORECASE)

        # 2. 查找所有命名参数 (例如 :案號, :銀行代碼)，并按顺序记录下来
        # re.findall会按从左到右的顺序找到所有匹配项
        named_params = re.findall(r':(\w+)', sql_query)
        
        # 3. 将所有命名参数替换为 '?'
        sql_query_fixed = re.sub(r':(\w+)', '?', sql_query)
        
        logger.debug(f"修复后的SQL: {sql_query_fixed}")
        logger.debug(f"解析出的参数顺序: {named_params}")
        
        return sql_query_fixed, named_params
    
    def _validate_and_fix_basic_sql(self, sql_query: str) -> str:
        """验证并修复基础SQL结构问题"""
        import re
        
        # 检查是否缺少FROM子句
        if not re.search(r'\bFROM\b', sql_query, re.IGNORECASE):
            logger.warning("SQL查询缺少FROM子句，尝试修复...")
            
            # 检查是否是查询字符串常数的无效SQL
            # 例如：SELECT '停话时间' as 停话时间, '开案编號' as 开案编號
            if re.search(r"SELECT\s+('[^']*'\s+as\s+\w+\s*,?\s*)+", sql_query, re.IGNORECASE):
                logger.warning("检测到无效的字符串常数查询，生成默认查询...")
                # 生成一个安全的默认查询
                return 'SELECT 1 as id, \'无数据\' as message'
        
        # 检查是否包含无效的表名引用
        if "'in'" in sql_query:
            # 'in'是SQL关键字，不应作为表名
            logger.warning("检测到使用SQL关键字'in'作为表名，进行修复...")
            sql_query = sql_query.replace("'in'", '"DataTable"')
        
        # 检查参数绑定是否合理
        question_marks = sql_query.count('?')
        where_clause = re.search(r'\bWHERE\b.*', sql_query, re.IGNORECASE)
        if where_clause and question_marks > 0:
            where_text = where_clause.group(0)
            # 检查是否有不合理的参数组合
            if question_marks % 2 != 0 and 'OR' in where_text and 'IS NULL' in where_text:
                logger.warning("检测到不平衡的可选参数绑定...")
        
        return sql_query
    
    def _add_parameters(self, root: ET.Element, parameters: List[Dict[str, Any]]):
        """添加报表参数 - 修复版本"""
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
        output_file = output_dir / f"{base_name}_fixed_generator.rptdesign"
        
        # 保存文件
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(xml_content)
        
        return str(output_file)

    def _add_custom_functions(self, root: ET.Element):
        """添加自定义JavaScript函数 - 修复版本"""
        # 在根元素下添加初始化方法
        existing_method = root.find('.//method[@name="initialize"]')
        if existing_method is not None:
            root.remove(existing_method)
        
        method = ET.Element('method')
        method.set('name', 'initialize')
        
        # JavaScript代码 - 保持原样，让ElementTree自动转义
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
        
        method.text = js_code  # ElementTree会自动转义，这是正确的
        
        # 插入到合适位置（在property元素之后）
        properties = root.findall('./property')
        if properties:
            insert_index = list(root).index(properties[-1]) + 1
            root.insert(insert_index, method)
        else:
            root.insert(0, method)

    def _add_professional_report_layout(self, root: ET.Element, analysis_result: ExcelAnalysisResult, ai_config: Dict[str, Any]):
        """添加专业的报表布局 - 修复版本"""
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
        """添加专业的报表标题区域 - 修复版本使用CDATA"""
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
        
        # >> 新的、更可靠的代码 <<
        # 1. 只插入一个简单、唯一的预留位置
        content.text = '__HTML_CONTENT_PLACEHOLDER__'

    def _add_date_info_section(self, body: ET.Element):
        """添加日期信息区域 - 修复版本"""
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
        # 直接写入表达式，XML库会自动转义 - 这是正确的行为
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
        # 直接写入表达式，让ElementTree自动转义
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


def test_generator():
    """测试生成器"""
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
        'sql_query': 'SELECT "TsUnit"."FCode" as 编码, "TsUnit"."FName" as 名称, "TsUnit"."FTable" as 类型 FROM "TsUnit" ORDER BY "TsUnit"."FCode"',
        'db_url': 'jdbc:jeedsql:jtds:sqlserver://localhost:1433;DatabaseName=QS;'
    }
    
    generator = BIRTGenerator()
    output_file = generator.generate_report(test_result, test_config, "output")
    print(f"修复版测试报表生成完成: {output_file}")


if __name__ == "__main__":
    test_generator()