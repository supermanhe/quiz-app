"""
调试工具 - 查看日志和数据库
"""
import os
import sqlite3
import sys

def view_log():
    """查看错误日志"""
    log_path = os.path.join(os.path.dirname(__file__), "debug.log")
    
    if not os.path.exists(log_path):
        print("📄 日志文件不存在，还没有记录过错误")
        return
    
    print("📄 ===== 错误日志 =====")
    try:
        with open(log_path, 'r', encoding='utf-8') as f:
            content = f.read()
            if content:
                print(content)
            else:
                print("日志文件为空")
    except Exception as e:
        print(f"读取日志失败: {e}")
    print("=====================")

def view_database(db_path=None):
    """查看数据库内容"""
    saves_dir = os.path.join(os.path.dirname(__file__), "saves")
    
    if db_path is None:
        # 列出所有存档
        if not os.path.exists(saves_dir):
            print("📁 saves 目录不存在")
            return
        
        db_files = [f for f in os.listdir(saves_dir) if f.endswith('.db')]
        if not db_files:
            print("📁 没有存档文件")
            return
        
        print("📁 可用存档文件:")
        for i, f in enumerate(db_files, 1):
            print(f"  {i}. {f}")
        
        choice = input("\n选择要查看的存档编号 (或输入完整路径): ").strip()
        
        if choice.isdigit() and 1 <= int(choice) <= len(db_files):
            db_path = os.path.join(saves_dir, db_files[int(choice)-1])
        elif os.path.exists(choice):
            db_path = choice
        else:
            print("❌ 无效选择")
            return
    
    print(f"\n📊 查看数据库: {db_path}")
    print("=" * 60)
    
    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        # 查看存档信息
        cursor.execute("SELECT * FROM save_info WHERE id = 1")
        save_info = cursor.fetchone()
        if save_info:
            print("\n📋 存档信息:")
            for key in save_info.keys():
                print(f"  {key}: {save_info[key]}")
        
        # 查看题目统计
        cursor.execute('''
            SELECT 
                COUNT(*) as total,
                SUM(CASE WHEN is_answered = 1 THEN 1 ELSE 0 END) as answered,
                SUM(CASE WHEN is_correct = 1 THEN 1 ELSE 0 END) as correct
            FROM questions
        ''')
        stats = cursor.fetchone()
        print(f"\n📈 答题统计:")
        print(f"  总题数: {stats['total']}")
        print(f"  已答题: {stats['answered']}")
        print(f"  正确数: {stats['correct']}")
        
        # 查看前5道题目
        print(f"\n📝 前5道题目预览:")
        cursor.execute("SELECT seq_no, q_type, question, answer, is_answered, is_correct FROM questions LIMIT 5")
        for row in cursor.fetchall():
            status = "✓" if row['is_answered'] and row['is_correct'] else "✗" if row['is_answered'] else "○"
            print(f"  [{status}] 第{row['seq_no']}题 [{row['q_type']}] {row['question'][:30]}... 答案:{row['answer']}")
        
        conn.close()
        
    except Exception as e:
        print(f"❌ 读取数据库失败: {e}")

def clear_log():
    """清空日志文件"""
    log_path = os.path.join(os.path.dirname(__file__), "debug.log")
    if os.path.exists(log_path):
        open(log_path, 'w').close()
        print("✅ 日志已清空")
    else:
        print("📄 日志文件不存在")

def main():
    print("=" * 60)
    print("  答题软件调试工具")
    print("=" * 60)
    
    while True:
        print("\n🔧 请选择操作:")
        print("  1. 查看错误日志")
        print("  2. 查看数据库内容")
        print("  3. 清空日志文件")
        print("  0. 退出")
        
        choice = input("\n选择: ").strip()
        
        if choice == '1':
            view_log()
        elif choice == '2':
            view_database()
        elif choice == '3':
            clear_log()
        elif choice == '0':
            break
        else:
            print("❌ 无效选择")
    
    input("\n按回车键退出...")

if __name__ == "__main__":
    main()
