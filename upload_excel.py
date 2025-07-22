#!/usr/bin/env python3
"""
Excel文件上传和处理工具
支持拖拽上传和批量处理
"""
import sys
import os
import shutil
from pathlib import Path
import argparse

# 添加src目录到Python路径
sys.path.append(str(Path(__file__).parent / "src"))

from src.main import BIRTAIGenerator


def copy_excel_files_to_input(file_paths):
    """将Excel文件复制到input目录"""
    input_dir = Path("input")
    input_dir.mkdir(exist_ok=True)
    
    copied_files = []
    
    for file_path in file_paths:
        source_path = Path(file_path)
        
        if not source_path.exists():
            print(f"❌ 文件不存在: {file_path}")
            continue
        
        if source_path.suffix.lower() not in ['.xlsx', '.xls']:
            print(f"⚠️  跳过非Excel文件: {source_path.name}")
            continue
        
        # 目标文件路径
        target_path = input_dir / source_path.name
        
        # 如果文件已存在，添加序号
        counter = 1
        original_target = target_path
        while target_path.exists():
            stem = original_target.stem
            suffix = original_target.suffix
            target_path = input_dir / f"{stem}_{counter}{suffix}"
            counter += 1
        
        try:
            shutil.copy2(source_path, target_path)
            copied_files.append(str(target_path))
            print(f"📁 已复制: {source_path.name} → {target_path.name}")
        except Exception as e:
            print(f"❌ 复制失败 {source_path.name}: {str(e)}")
    
    return copied_files


def interactive_file_upload():
    """交互式文件上传"""
    print("📤 Excel文件上传工具")
    print("=" * 50)
    
    print("📋 上传方式:")
    print("1. 输入文件路径")
    print("2. 输入目录路径（批量上传）")
    print("3. 拖拽文件到终端（Windows）")
    
    while True:
        user_input = input("\n请输入Excel文件或目录路径（'q'退出）: ").strip()
        
        if user_input.lower() == 'q':
            break
        
        if not user_input:
            continue
        
        # 移除引号（拖拽文件时可能包含引号）
        user_input = user_input.strip('"\'')
        
        input_path = Path(user_input)
        
        if not input_path.exists():
            print(f"❌ 路径不存在: {user_input}")
            continue
        
        file_paths = []
        
        if input_path.is_file():
            file_paths = [str(input_path)]
        elif input_path.is_dir():
            # 查找目录中的Excel文件
            excel_files = list(input_path.glob("*.xlsx")) + list(input_path.glob("*.xls"))
            file_paths = [str(f) for f in excel_files]
            
            if not file_paths:
                print(f"⚠️  目录中未找到Excel文件: {input_path}")
                continue
            
            print(f"🔍 找到 {len(file_paths)} 个Excel文件")
        
        # 复制文件到input目录
        copied_files = copy_excel_files_to_input(file_paths)
        
        if copied_files:
            print(f"✅ 成功上传 {len(copied_files)} 个文件")
            
            # 询问是否立即处理
            process_now = input("是否立即处理这些文件? (Y/n): ").strip().lower()
            if process_now != 'n':
                process_uploaded_files(copied_files)
        else:
            print("❌ 没有成功上传任何文件")


def process_uploaded_files(file_paths=None):
    """处理已上传的文件"""
    print("\n🚀 开始处理Excel文件...")
    
    try:
        # 初始化BIRT生成器
        generator = BIRTAIGenerator()
        
        if file_paths:
            # 处理指定文件
            results = []
            for file_path in file_paths:
                print(f"\n{'='*20} 处理: {Path(file_path).name} {'='*20}")
                result = generator.process_single_file(file_path)
                results.append(result)
                
                if result['status'] == 'success':
                    print(f"✅ 成功: {result['output_file']}")
                else:
                    print(f"❌ 失败: {result['error']}")
        else:
            # 批量处理input目录中的所有文件
            summary = generator.process_batch("input", "output")
            print(f"\n📊 批量处理完成:")
            print(f"  总计: {summary['total_files']} 个文件")
            print(f"  成功: {summary['successful']} 个")
            print(f"  失败: {summary['failed']} 个")
            print(f"  成功率: {summary['success_rate']:.1%}")
    
    except Exception as e:
        print(f"❌ 处理失败: {str(e)}")


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="Excel文件上传和处理工具")
    parser.add_argument('files', nargs='*', help='要上传的Excel文件路径')
    parser.add_argument('--interactive', '-i', action='store_true', help='交互式上传模式')
    parser.add_argument('--process-only', '-p', action='store_true', help='只处理已上传的文件')
    parser.add_argument('--output', '-o', default='output', help='输出目录')
    
    args = parser.parse_args()
    
    print("📊 Excel文件上传和处理工具")
    print("=" * 50)
    
    try:
        if args.process_only:
            # 只处理已有文件
            process_uploaded_files()
        elif args.files:
            # 命令行指定文件
            copied_files = copy_excel_files_to_input(args.files)
            if copied_files:
                process_uploaded_files(copied_files)
        elif args.interactive:
            # 交互式模式
            interactive_file_upload()
        else:
            # 默认：显示使用说明
            print("📖 使用方法:")
            print("1. 上传并处理文件:")
            print("   python upload_excel.py file1.xlsx file2.xlsx")
            print("   python upload_excel.py C:\\path\\to\\excel\\directory")
            print("")
            print("2. 交互式上传:")
            print("   python upload_excel.py --interactive")
            print("")
            print("3. 只处理已上传的文件:")
            print("   python upload_excel.py --process-only")
            print("")
            print("4. 直接拖拽文件:")
            print("   将Excel文件拖拽到此程序图标上")
            
            # 检查是否有待处理的文件
            input_dir = Path("input")
            if input_dir.exists():
                excel_files = list(input_dir.glob("*.xlsx")) + list(input_dir.glob("*.xls"))
                if excel_files:
                    print(f"\n📁 发现 {len(excel_files)} 个待处理文件:")
                    for file in excel_files:
                        print(f"  📄 {file.name}")
                    
                    process_now = input("\n是否处理这些文件? (Y/n): ").strip().lower()
                    if process_now != 'n':
                        process_uploaded_files()
        
        print("\n🎉 操作完成!")
        print(f"📁 生成的BIRT文件位于: {args.output}/ 目录")
        
        return 0
        
    except KeyboardInterrupt:
        print("\n\n🚫 用户取消操作")
        return 1
    except Exception as e:
        print(f"\n❌ 操作失败: {str(e)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())