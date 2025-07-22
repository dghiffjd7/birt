#!/usr/bin/env python3
"""
BIRT .rptdesign文件验证工具
检查生成的报表文件是否符合BIRT标准，能否在Eclipse BIRT中正确打开
"""

import xml.etree.ElementTree as ET
from pathlib import Path
import sys
from typing import Dict, List, Tuple
import re

class BIRTValidator:
    """BIRT报表文件验证器"""
    
    def __init__(self):
        self.namespace = {"": "http://www.eclipse.org/birt/2005/design"}
        self.errors = []
        self.warnings = []
        
    def validate_file(self, file_path: str) -> Dict[str, any]:
        """
        验证单个rptdesign文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            验证结果字典
        """
        self.errors = []
        self.warnings = []
        
        print(f"🔍 验证文件: {file_path}")
        
        try:
            # 1. 检查文件存在性
            if not Path(file_path).exists():
                self.errors.append("文件不存在")
                return self._create_result(False)
            
            # 2. 读取文件内容
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 3. XML格式验证
            if not self._validate_xml_format(content):
                return self._create_result(False)
            
            # 4. 解析XML
            root = ET.fromstring(content)
            
            # 5. BIRT结构验证
            self._validate_birt_structure(root, content)
            
            # 6. Quicksilver特定验证
            self._validate_quicksilver_integration(root, content)
            
            # 7. 数据质量验证
            self._validate_data_quality(root, content)
            
            success = len(self.errors) == 0
            return self._create_result(success)
            
        except ET.ParseError as e:
            self.errors.append(f"XML解析错误: {str(e)}")
            return self._create_result(False)
        except Exception as e:
            self.errors.append(f"验证过程错误: {str(e)}")
            return self._create_result(False)
    
    def _validate_xml_format(self, content: str) -> bool:
        """验证XML格式"""
        # 检查XML声明
        if not content.strip().startswith('<?xml'):
            self.errors.append("缺少XML声明")
            return False
        
        # 检查编码
        if 'encoding="UTF-8"' not in content:
            self.warnings.append("建议使用UTF-8编码")
        
        # 检查格式是否美化
        if '<body><label id=' in content:  # 检查是否压缩格式
            self.warnings.append("XML格式未美化，建议使用格式化的XML以便调试")
        
        return True
    
    def _validate_birt_structure(self, root: ET.Element, content: str):
        """验证BIRT结构"""
        # 检查根元素 - 考虑命名空间
        root_tag = root.tag
        if root_tag.endswith('}report'):
            root_tag = 'report'  # 移除命名空间前缀
        
        if root_tag != 'report':
            self.errors.append(f"根元素必须是'report'，当前是'{root_tag}'")
        
        # 检查命名空间
        if 'http://www.eclipse.org/birt/2005/design' not in content:
            self.errors.append("缺少BIRT命名空间")
        
        # 检查版本
        version = root.get('version')
        if not version:
            self.warnings.append("缺少版本信息")
        elif version not in ['3.2.23', '4.10.0']:
            self.warnings.append(f"版本 {version} 可能不兼容")
        
        # 检查必需的属性
        required_properties = ['createdBy', 'units', 'bidiLayoutOrientation']
        for prop in required_properties:
            # 使用更宽松的查找方式
            prop_elem = root.find(f'.//*[@name="{prop}"]')
            if prop_elem is None:
                self.warnings.append(f"缺少推荐属性: {prop}")
        
        # 检查数据源 - 使用更宽松的查找
        datasources = root.findall('.//*')
        datasource_found = any('oda-data-source' in elem.tag for elem in datasources)
        if not datasource_found:
            self.errors.append("缺少数据源配置")
        
        # 检查数据集
        dataset_found = any('oda-data-set' in elem.tag for elem in datasources)
        if not dataset_found:
            self.errors.append("缺少数据集配置")
        
        # 检查页面设置
        page_setup_found = any('page-setup' in elem.tag or 'simple-master-page' in elem.tag for elem in datasources)
        if not page_setup_found:
            self.warnings.append("缺少页面设置")
        
        # 检查报表主体
        body_found = any('body' in elem.tag for elem in datasources) or '<body>' in content
        if not body_found:
            self.errors.append("缺少报表主体")
        
        print("✅ BIRT结构验证完成")
    
    def _validate_quicksilver_integration(self, root: ET.Element, content: str):
        """验证Quicksilver集成"""
        # 检查驱动配置 - 直接在内容中搜索
        if 'com.jeedsoft.jeedsql.jdbc.Driver' not in content:
            self.errors.append("缺少JeedSQL数据库驱动配置")
        
        # 检查初始化脚本
        if 'ReportUtil.initializeDataSource' not in content:
            self.errors.append("缺少Quicksilver初始化脚本")
        
        if 'importPackage(Packages.com.jeedsoft.quicksilver.report.util)' not in content:
            self.errors.append("缺少Quicksilver包导入")
        
        # 检查数据库URL格式
        if 'jdbc:jeedsql' not in content:
            self.warnings.append("数据库URL可能不是JeedSQL格式")
        
        print("✅ Quicksilver集成验证完成")
    
    def _validate_data_quality(self, root: ET.Element, content: str):
        """验证数据质量"""
        # 检查SQL查询 - 在内容中搜索
        import re
        query_pattern = r'<property name="queryText"[^>]*>(.*?)</property>'
        queries = re.findall(query_pattern, content, re.DOTALL)
        
        if not queries:
            self.warnings.append("未找到SQL查询配置")
        
        for sql in queries:
            # 清理SQL内容
            sql = sql.strip()
            if not sql:
                continue
            
            # 检查字段名引号
            if "'" in sql and '"' not in sql:
                self.warnings.append("建议使用双引号包围字段名")
            
            # 检查中文字段处理
            chinese_pattern = r'[\u4e00-\u9fff]'
            if re.search(chinese_pattern, sql):
                if not re.search(r'"[^"]*[\u4e00-\u9fff][^"]*"', sql):
                    self.warnings.append("中文字段名应使用双引号包围")
        
        # 检查参数配置
        param_count = content.count('scalar-parameter')
        binding_count = content.count('parameterBindings')
        
        if param_count > 0 and binding_count == 0:
            self.warnings.append("有参数定义但缺少参数绑定")
        
        # 检查结果集配置
        if 'resultSet' not in content and 'result-set-column' not in content:
            self.warnings.append("缺少结果集结构定义")
        
        print("✅ 数据质量验证完成")
    
    def _create_result(self, success: bool) -> Dict[str, any]:
        """创建验证结果"""
        return {
            'success': success,
            'errors': self.errors.copy(),
            'warnings': self.warnings.copy(),
            'total_issues': len(self.errors) + len(self.warnings)
        }
    
    def print_result(self, result: Dict[str, any], file_path: str):
        """打印验证结果"""
        print("\n" + "="*60)
        print(f"📊 验证结果: {file_path}")
        print("="*60)
        
        if result['success']:
            print("✅ 验证通过 - 文件可以在Eclipse BIRT中正常打开")
        else:
            print("❌ 验证失败 - 文件可能无法在Eclipse BIRT中正常工作")
        
        if result['errors']:
            print(f"\n🚨 错误 ({len(result['errors'])}个):")
            for i, error in enumerate(result['errors'], 1):
                print(f"  {i}. {error}")
        
        if result['warnings']:
            print(f"\n⚠️  警告 ({len(result['warnings'])}个):")
            for i, warning in enumerate(result['warnings'], 1):
                print(f"  {i}. {warning}")
        
        if not result['errors'] and not result['warnings']:
            print("\n🎉 完美！没有发现任何问题")
        
        print("\n" + "="*60)


def main():
    """主函数"""
    import argparse
    
    parser = argparse.ArgumentParser(description='验证BIRT rptdesign文件')
    parser.add_argument('file', nargs='?', help='要验证的rptdesign文件路径')
    parser.add_argument('--all', action='store_true', help='验证output目录中的所有rptdesign文件')
    parser.add_argument('--output-dir', default='output', help='输出目录路径 (默认: output)')
    
    args = parser.parse_args()
    
    validator = BIRTValidator()
    
    if args.all:
        # 验证所有文件
        output_dir = Path(args.output_dir)
        rptdesign_files = list(output_dir.glob("*.rptdesign"))
        
        if not rptdesign_files:
            print(f"❌ 在 {output_dir} 目录中未找到.rptdesign文件")
            return
        
        print(f"🔍 找到 {len(rptdesign_files)} 个.rptdesign文件")
        
        all_results = []
        for file_path in rptdesign_files:
            result = validator.validate_file(str(file_path))
            validator.print_result(result, str(file_path))
            all_results.append((file_path.name, result))
        
        # 总结
        print("\n" + "="*60)
        print("📈 总体验证结果")
        print("="*60)
        
        success_count = sum(1 for _, result in all_results if result['success'])
        total_files = len(all_results)
        
        print(f"✅ 通过验证: {success_count}/{total_files}")
        print(f"❌ 验证失败: {total_files - success_count}/{total_files}")
        
        if success_count == total_files:
            print("\n🎉 所有文件都可以在Eclipse BIRT中正常使用！")
        else:
            print("\n⚠️  部分文件需要修复才能在Eclipse BIRT中正常使用")
    
    elif args.file:
        # 验证单个文件
        result = validator.validate_file(args.file)
        validator.print_result(result, args.file)
        
        if result['success']:
            print("\n💡 建议:")
            print("  1. 在Eclipse BIRT Designer中打开此文件")
            print("  2. 检查数据源连接配置")
            print("  3. 预览报表确认显示效果")
        else:
            print("\n🔧 修复建议:")
            print("  1. 运行 python fix_rptdesign.py 修复格式问题")
            print("  2. 检查数据库连接配置")
            print("  3. 重新生成报表文件")
    
    else:
        parser.print_help()


if __name__ == "__main__":
    main() 