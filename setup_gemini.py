#!/usr/bin/env python3
"""
Gemini API配置助手
帮助用户配置Gemini API密钥
"""
import os
import getpass
from pathlib import Path
from dotenv import load_dotenv, set_key


def setup_gemini_api():
    """设置Gemini API密钥"""
    print("🚀 Gemini API配置助手")
    print("=" * 50)
    
    # 检查.env文件
    env_file = Path(".env")
    if not env_file.exists():
        print("📄 创建.env配置文件...")
        # 复制示例文件
        example_file = Path(".env.example")
        if example_file.exists():
            env_file.write_text(example_file.read_text(encoding='utf-8'), encoding='utf-8')
        else:
            env_file.write_text("# Gemini API配置\nGEMINI_API_KEY=\n", encoding='utf-8')
    
    # 加载现有配置
    load_dotenv(env_file)
    current_key = os.getenv("GEMINI_API_KEY")
    
    if current_key:
        print(f"✅ 当前已配置Gemini API密钥: {current_key[:10]}...")
        
        choice = input("\n是否要更新API密钥? (y/N): ").strip().lower()
        if choice != 'y':
            print("🎉 保持现有配置")
            return True
    
    print("\n🔑 配置Gemini API密钥")
    print("请访问以下网址获取API密钥:")
    print("🌐 https://makersuite.google.com/app/apikey")
    print("\n步骤:")
    print("1. 登录Google账号")
    print("2. 点击 'Create API Key'")
    print("3. 选择项目或创建新项目")
    print("4. 复制生成的API密钥")
    
    # 获取API密钥
    while True:
        api_key = getpass.getpass("\n请输入您的Gemini API密钥: ").strip()
        
        if not api_key:
            print("❌ API密钥不能为空")
            continue
        
        if len(api_key) < 20:
            print("❌ API密钥长度似乎不正确，请检查")
            continue
        
        break
    
    # 测试API密钥
    print("\n🧪 测试API密钥...")
    try:
        import google.generativeai as genai
        genai.configure(api_key=api_key)
        
        # 简单测试
        model = genai.GenerativeModel("gemini-2.5-pro")
        response = model.generate_content("Hello, this is a test.")
        
        if response.text:
            print("✅ API密钥测试成功!")
        else:
            print("⚠️  API调用成功但响应为空，可能是配额问题")
            
    except Exception as e:
        print(f"❌ API密钥测试失败: {str(e)}")
        
        choice = input("是否仍要保存此API密钥? (y/N): ").strip().lower()
        if choice != 'y':
            print("🚫 取消配置")
            return False
    
    # 保存API密钥
    try:
        set_key(env_file, "GEMINI_API_KEY", api_key)
        print(f"💾 API密钥已保存到 {env_file}")
        return True
        
    except Exception as e:
        print(f"❌ 保存配置失败: {str(e)}")
        return False


def verify_setup():
    """验证配置"""
    print("\n🔍 验证Gemini配置...")
    
    # 检查依赖
    try:
        import google.generativeai as genai
        print("✅ google-generativeai 库已安装")
    except ImportError:
        print("❌ 缺少 google-generativeai 库")
        print("请运行: pip install google-generativeai")
        return False
    
    # 检查API密钥
    load_dotenv()
    api_key = os.getenv("GEMINI_API_KEY")
    
    if not api_key:
        print("❌ 未找到GEMINI_API_KEY")
        return False
    
    print(f"✅ 找到API密钥: {api_key[:10]}...")
    
    # 测试连接
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel("gemini-2.5-pro")
        response = model.generate_content("Test connection")
        print("✅ Gemini API连接正常")
        return True
        
    except Exception as e:
        print(f"❌ Gemini API连接失败: {str(e)}")
        return False


def main():
    """主函数"""
    try:
        if setup_gemini_api():
            if verify_setup():
                print("\n🎉 Gemini配置完成!")
                print("\n📖 下一步:")
                print("1. 运行示例: python run_example.py")
                print("2. 处理Excel文件: python src/main.py -i input/your_file.xlsx")
                return 0
            else:
                print("\n⚠️  配置验证失败，请检查API密钥")
                return 1
        else:
            print("\n🚫 配置取消")
            return 1
            
    except KeyboardInterrupt:
        print("\n\n🚫 用户取消操作")
        return 1
    except Exception as e:
        print(f"\n❌ 配置失败: {str(e)}")
        return 1


if __name__ == "__main__":
    import sys
    sys.exit(main())