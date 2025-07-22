"""
BIRT AI生成器主程序
批量处理Excel文件并生成BIRT报表
"""
import asyncio
import logging
import sys
from pathlib import Path
from typing import List, Dict, Any
import json
from datetime import datetime
import argparse

# 添加src目录到Python路径
sys.path.append(str(Path(__file__).parent))

from analyzers.excel_analyzer import ExcelAnalyzer
from generators.birt_generator import BIRTGenerator
from utils.gemini_analyzer import SimpleAIAnalyzer
from utils.config import Config
from utils.logger import setup_logger


class BIRTAIGenerator:
    """BIRT AI生成器主类"""
    
    def __init__(self, config_path: str = ".env"):
        self.config = Config(config_path)
        self.logger = setup_logger(self.config.log_level)
        
        # 初始化组件
        self.excel_analyzer = ExcelAnalyzer()
        self.birt_generator = BIRTGenerator()
        self.ai_analyzer = SimpleAIAnalyzer(self.config.gemini_api_key)
        
        self.logger.info("BIRT AI生成器初始化完成")
    
    def process_single_file(self, excel_path: str, output_dir: str = "output") -> Dict[str, Any]:
        """
        处理单个Excel文件
        
        Args:
            excel_path: Excel文件路径
            output_dir: 输出目录
            
        Returns:
            Dict: 处理结果
        """
        start_time = datetime.now()
        file_path = Path(excel_path)
        
        self.logger.info(f"开始处理文件: {file_path.name}")
        
        try:
            # 1. 分析Excel文件
            self.logger.info("步骤1: 分析Excel文件结构")
            excel_result = self.excel_analyzer.analyze_excel(str(file_path))
            
            # 2. AI分析生成配置
            self.logger.info("步骤2: AI分析生成BIRT配置")
            ai_config = self.ai_analyzer.analyze_excel_simple(excel_result)
            
            # 3. 生成BIRT报表
            self.logger.info("步骤3: 生成BIRT报表文件")
            output_file = self.birt_generator.generate_report(
                excel_result, 
                ai_config, 
                output_dir
            )
            
            end_time = datetime.now()
            processing_time = (end_time - start_time).total_seconds()
            
            result = {
                'status': 'success',
                'input_file': str(file_path),
                'output_file': output_file,
                'processing_time': processing_time,
                'excel_analysis': excel_result.to_dict(),
                'ai_config': ai_config,
                'timestamp': end_time.isoformat()
            }
            
            self.logger.info(f"文件处理完成: {file_path.name}, 耗时: {processing_time:.2f}秒")
            return result
            
        except Exception as e:
            end_time = datetime.now()
            processing_time = (end_time - start_time).total_seconds()
            
            error_result = {
                'status': 'failed',
                'input_file': str(file_path),
                'error': str(e),
                'processing_time': processing_time,
                'timestamp': end_time.isoformat()
            }
            
            self.logger.error(f"处理文件失败: {file_path.name}, 错误: {str(e)}")
            return error_result
    
    def process_batch(self, input_dir: str, output_dir: str = "output", pattern: str = "*.xlsx") -> Dict[str, Any]:
        """
        批量处理Excel文件
        
        Args:
            input_dir: 输入目录
            output_dir: 输出目录  
            pattern: 文件匹配模式
            
        Returns:
            Dict: 批量处理结果
        """
        input_path = Path(input_dir)
        if not input_path.exists():
            raise FileNotFoundError(f"输入目录不存在: {input_dir}")
        
        # 查找Excel文件
        excel_files = list(input_path.glob(pattern))
        if not excel_files:
            self.logger.warning(f"在目录 {input_dir} 中未找到匹配的Excel文件")
            return {
                'status': 'completed',
                'total_files': 0,
                'successful': 0,
                'failed': 0,
                'results': []
            }
        
        self.logger.info(f"找到 {len(excel_files)} 个Excel文件，开始批量处理")
        
        results = []
        successful = 0
        failed = 0
        
        for excel_file in excel_files:
            result = self.process_single_file(str(excel_file), output_dir)
            results.append(result)
            
            if result['status'] == 'success':
                successful += 1
            else:
                failed += 1
        
        # 生成汇总报告
        summary = {
            'status': 'completed',
            'total_files': len(excel_files),
            'successful': successful,
            'failed': failed,
            'success_rate': successful / len(excel_files) if excel_files else 0,
            'results': results,
            'timestamp': datetime.now().isoformat()
        }
        
        # 保存处理报告
        self._save_batch_report(summary, output_dir)
        
        self.logger.info(f"批量处理完成: 总计{len(excel_files)}个文件, 成功{successful}个, 失败{failed}个")
        return summary
    
    def _save_batch_report(self, summary: Dict[str, Any], output_dir: str):
        """保存批量处理报告"""
        try:
            output_path = Path(output_dir)
            output_path.mkdir(parents=True, exist_ok=True)
            
            report_file = output_path / f"batch_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            
            with open(report_file, 'w', encoding='utf-8') as f:
                json.dump(summary, f, ensure_ascii=False, indent=2)
            
            self.logger.info(f"批量处理报告已保存: {report_file}")
            
        except Exception as e:
            self.logger.error(f"保存批量处理报告失败: {str(e)}")
    
    def list_input_files(self, input_dir: str, pattern: str = "*.xlsx") -> List[str]:
        """列出输入目录中的Excel文件"""
        input_path = Path(input_dir)
        if not input_path.exists():
            return []
        
        excel_files = list(input_path.glob(pattern))
        return [str(f) for f in excel_files]
    
    def validate_environment(self) -> Dict[str, bool]:
        """验证运行环境"""
        checks = {}
        
        # 检查必要的目录
        checks['input_dir_exists'] = Path("input").exists()
        checks['output_dir_writable'] = True
        
        try:
            test_output = Path("output")
            test_output.mkdir(exist_ok=True)
            test_file = test_output / "test_write.tmp"
            test_file.write_text("test")
            test_file.unlink()
        except:
            checks['output_dir_writable'] = False
        
        # 检查模板文件
        checks['templates_exist'] = (Path("templates/birt/base_template.xml").exists())
        
        # 检查配置
        checks['config_valid'] = bool(self.config.gemini_api_key) if hasattr(self.config, 'gemini_api_key') else False
        
        return checks


def main():
    """命令行主程序"""
    parser = argparse.ArgumentParser(description="BIRT AI报表生成器")
    parser.add_argument('--input', '-i', required=True, help='输入Excel文件或目录')
    parser.add_argument('--output', '-o', default='output', help='输出目录 (默认: output)')
    parser.add_argument('--pattern', '-p', default='*.xlsx', help='文件匹配模式 (默认: *.xlsx)')
    parser.add_argument('--config', '-c', default='.env', help='配置文件路径 (默认: .env)')
    parser.add_argument('--validate', action='store_true', help='验证环境配置')
    parser.add_argument('--list', action='store_true', help='只列出匹配的文件，不处理')
    
    args = parser.parse_args()
    
    try:
        # 初始化生成器
        generator = BIRTAIGenerator(args.config)
        
        # 验证环境
        if args.validate:
            checks = generator.validate_environment()
            print("环境检查结果:")
            for check, passed in checks.items():
                status = "✓" if passed else "✗"
                print(f"  {status} {check}")
            
            if not all(checks.values()):
                print("\n某些检查未通过，请检查配置")
                return 1
            else:
                print("\n环境检查通过")
                return 0
        
        input_path = Path(args.input)
        
        # 列出文件
        if args.list:
            if input_path.is_file():
                files = [str(input_path)]
            else:
                files = generator.list_input_files(str(input_path), args.pattern)
            
            print(f"找到 {len(files)} 个匹配的文件:")
            for f in files:
                print(f"  {f}")
            return 0
        
        # 处理文件
        if input_path.is_file():
            # 单文件处理
            result = generator.process_single_file(str(input_path), args.output)
            if result['status'] == 'success':
                print(f"处理成功: {result['output_file']}")
                return 0
            else:
                print(f"处理失败: {result['error']}")
                return 1
        else:
            # 批量处理
            summary = generator.process_batch(str(input_path), args.output, args.pattern)
            print(f"批量处理完成:")
            print(f"  总计: {summary['total_files']} 个文件")
            print(f"  成功: {summary['successful']} 个")
            print(f"  失败: {summary['failed']} 个")
            print(f"  成功率: {summary['success_rate']:.1%}")
            
            return 0 if summary['failed'] == 0 else 1
    
    except Exception as e:
        print(f"程序执行出错: {str(e)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())