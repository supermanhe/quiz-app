"""
数据库管理模块 - 处理SQLite存档的创建、读取、更新、删除
"""
import sqlite3
import os
import json
from datetime import datetime
from typing import List, Dict, Optional, Any

class QuizDatabase:
    """答题软件数据库管理类"""
    
    def __init__(self, db_path: str):
        """
        初始化数据库连接
        
        Args:
            db_path: 数据库文件路径
        """
        self.db_path = db_path
        self.conn = None
        self.cursor = None
        self.connect()
        self.init_tables()
    
    def connect(self):
        """建立数据库连接"""
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row
        self.cursor = self.conn.cursor()
    
    def close(self):
        """关闭数据库连接"""
        if self.conn:
            self.conn.close()
            self.conn = None
            self.cursor = None
    
    def init_tables(self):
        """初始化数据库表结构"""
        # 题目表 - 存储所有题目和答题状态
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS questions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                seq_no INTEGER UNIQUE,           -- 序号
                q_type TEXT NOT NULL,            -- 题型：single/multiple/judge
                question TEXT NOT NULL,          -- 题干
                option_a TEXT,                   -- 选项A
                option_b TEXT,                   -- 选项B
                option_c TEXT,                   -- 选项C
                option_d TEXT,                   -- 选项D
                option_e TEXT,                   -- 选项E
                option_f TEXT,                   -- 选项F
                answer TEXT NOT NULL,            -- 正确答案
                wrong_count INTEGER DEFAULT 0,   -- 错题次数
                time_spent REAL DEFAULT 0,       -- 做题时间（秒）
                is_favorite INTEGER DEFAULT 0,   -- 是否收藏 0/1
                is_answered INTEGER DEFAULT 0,   -- 是否已做 0/1
                user_answer TEXT,                -- 用户答案
                is_correct INTEGER DEFAULT 0,    -- 是否正确 0/1
                first_answer_time TEXT,          -- 首次答题时间
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # 存档信息表
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS save_info (
                id INTEGER PRIMARY KEY CHECK (id = 1),
                save_name TEXT NOT NULL,         -- 存档名称
                total_count INTEGER DEFAULT 0,   -- 总题数
                answered_count INTEGER DEFAULT 0, -- 已答题数
                correct_count INTEGER DEFAULT 0, -- 正确数
                wrong_count INTEGER DEFAULT 0,   -- 错误数
                favorite_count INTEGER DEFAULT 0, -- 收藏数
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        self.conn.commit()
    
    def import_questions(self, questions: List[Dict]):
        """
        导入题目到数据库
        
        Args:
            questions: 题目列表，每项为包含题目信息的字典
        """
        self.cursor.execute("DELETE FROM questions")
        
        for q in questions:
            self.cursor.execute('''
                INSERT INTO questions (
                    seq_no, q_type, question, 
                    option_a, option_b, option_c, option_d, option_e, option_f,
                    answer
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                q.get('seq_no'),
                q.get('type'),
                q.get('question'),
                q.get('option_a', ''),
                q.get('option_b', ''),
                q.get('option_c', ''),
                q.get('option_d', ''),
                q.get('option_e', ''),
                q.get('option_f', ''),
                q.get('answer', '')
            ))
        
        # 更新存档统计信息
        self._update_save_stats()
        self.conn.commit()
    
    def get_all_questions(self) -> List[Dict]:
        """获取所有题目"""
        self.cursor.execute('SELECT * FROM questions ORDER BY seq_no')
        rows = self.cursor.fetchall()
        return [dict(row) for row in rows]
    
    def get_question_by_seq(self, seq_no: int) -> Optional[Dict]:
        """根据序号获取题目"""
        self.cursor.execute('SELECT * FROM questions WHERE seq_no = ?', (seq_no,))
        row = self.cursor.fetchone()
        return dict(row) if row else None
    
    def get_question_by_id(self, qid: int) -> Optional[Dict]:
        """根据ID获取题目"""
        self.cursor.execute('SELECT * FROM questions WHERE id = ?', (qid,))
        row = self.cursor.fetchone()
        return dict(row) if row else None
    
    def update_answer(self, seq_no: int, user_answer: str, is_correct: bool, 
                      time_spent: float, wrong_increment: int = 0):
        """
        更新答题结果
        
        Args:
            seq_no: 题目序号
            user_answer: 用户答案
            is_correct: 是否正确
            time_spent: 用时（秒）
            wrong_increment: 错题次数增量
        """
        # 获取当前记录
        self.cursor.execute(
            'SELECT wrong_count, time_spent, first_answer_time FROM questions WHERE seq_no = ?',
            (seq_no,)
        )
        row = self.cursor.fetchone()
        
        if row:
            current_wrong = row['wrong_count']
            current_time = row['time_spent'] or 0
            first_time = row['first_answer_time']
            
            # 只记录首次答题时间
            if not first_time:
                first_time = datetime.now().isoformat()
            
            # 如果之前没有答过，记录时间；否则保留原时间
            new_time = time_spent if current_time == 0 else current_time
            
            self.cursor.execute('''
                UPDATE questions SET
                    user_answer = ?,
                    is_correct = ?,
                    is_answered = 1,
                    wrong_count = wrong_count + ?,
                    time_spent = ?,
                    first_answer_time = ?
                WHERE seq_no = ?
            ''', (user_answer, 1 if is_correct else 0, wrong_increment,
                  new_time, first_time, seq_no))
            
            self._update_save_stats()
            self.conn.commit()
    
    def toggle_favorite(self, seq_no: int) -> bool:
        """
        切换收藏状态
        
        Returns:
            切换后的收藏状态
        """
        self.cursor.execute(
            'SELECT is_favorite FROM questions WHERE seq_no = ?', (seq_no,)
        )
        row = self.cursor.fetchone()
        
        if row:
            new_status = 0 if row['is_favorite'] else 1
            self.cursor.execute(
                'UPDATE questions SET is_favorite = ? WHERE seq_no = ?',
                (new_status, seq_no)
            )
            self._update_save_stats()
            self.conn.commit()
            return new_status == 1
        return False
    
    def get_statistics(self) -> Dict[str, Any]:
        """获取答题统计信息"""
        self.cursor.execute('''
            SELECT 
                COUNT(*) as total,
                SUM(CASE WHEN is_answered = 1 THEN 1 ELSE 0 END) as answered,
                SUM(CASE WHEN is_correct = 1 THEN 1 ELSE 0 END) as correct,
                SUM(CASE WHEN is_correct = 0 AND is_answered = 1 THEN 1 ELSE 0 END) as wrong,
                SUM(CASE WHEN is_favorite = 1 THEN 1 ELSE 0 END) as favorite
            FROM questions
        ''')
        row = self.cursor.fetchone()
        
        stats = {
            'total': row['total'] or 0,
            'answered': row['answered'] or 0,
            'correct': row['correct'] or 0,
            'wrong': row['wrong'] or 0,
            'favorite': row['favorite'] or 0
        }
        
        stats['accuracy'] = (stats['correct'] / stats['answered'] * 100) if stats['answered'] > 0 else 0
        
        return stats
    
    def _update_save_stats(self):
        """更新存档统计信息"""
        stats = self.get_statistics()
        
        self.cursor.execute('SELECT id FROM save_info WHERE id = 1')
        if self.cursor.fetchone():
            self.cursor.execute('''
                UPDATE save_info SET
                    total_count = ?,
                    answered_count = ?,
                    correct_count = ?,
                    wrong_count = ?,
                    favorite_count = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = 1
            ''', (stats['total'], stats['answered'], stats['correct'],
                  stats['wrong'], stats['favorite']))
        else:
            self.cursor.execute('''
                INSERT INTO save_info (id, save_name, total_count, answered_count,
                    correct_count, wrong_count, favorite_count)
                VALUES (1, '默认存档', ?, ?, ?, ?, ?)
            ''', (stats['total'], stats['answered'], stats['correct'],
                  stats['wrong'], stats['favorite']))
    
    def get_answered_questions(self) -> List[Dict]:
        """获取所有已答题目"""
        self.cursor.execute(
            'SELECT * FROM questions WHERE is_answered = 1 ORDER BY seq_no'
        )
        return [dict(row) for row in self.cursor.fetchall()]
    
    def get_wrong_questions(self) -> List[Dict]:
        """获取错题（错题次数>0）"""
        self.cursor.execute(
            'SELECT * FROM questions WHERE wrong_count > 0 ORDER BY seq_no'
        )
        return [dict(row) for row in self.cursor.fetchall()]
    
    def get_favorite_questions(self) -> List[Dict]:
        """获取收藏题目"""
        self.cursor.execute(
            'SELECT * FROM questions WHERE is_favorite = 1 ORDER BY seq_no'
        )
        return [dict(row) for row in self.cursor.fetchall()]
    
    def get_timeout_questions(self, timeout_seconds: float) -> List[Dict]:
        """获取超时题目"""
        self.cursor.execute(
            'SELECT * FROM questions WHERE is_answered = 1 AND time_spent > ? ORDER BY seq_no',
            (timeout_seconds,)
        )
        return [dict(row) for row in self.cursor.fetchall()]
    
    def update_save_name(self, name: str):
        """更新存档名称"""
        self.cursor.execute(
            'UPDATE save_info SET save_name = ? WHERE id = 1', (name,)
        )
        self.conn.commit()
    
    def get_save_name(self) -> str:
        """获取存档名称"""
        self.cursor.execute('SELECT save_name FROM save_info WHERE id = 1')
        row = self.cursor.fetchone()
        return row['save_name'] if row else '默认存档'


def list_saves(save_dir: str) -> List[Dict]:
    """
    列出所有存档
    
    Args:
        save_dir: 存档目录
        
    Returns:
        存档信息列表
    """
    saves = []
    if not os.path.exists(save_dir):
        return saves
    
    for filename in os.listdir(save_dir):
        if filename.endswith('.db'):
            db_path = os.path.join(save_dir, filename)
            try:
                conn = sqlite3.connect(db_path)
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT save_name, total_count, answered_count, 
                           correct_count, updated_at 
                    FROM save_info WHERE id = 1
                ''')
                row = cursor.fetchone()
                if row:
                    saves.append({
                        'filename': filename,
                        'name': row['save_name'],
                        'total': row['total_count'],
                        'answered': row['answered_count'],
                        'correct': row['correct_count'],
                        'updated': row['updated_at']
                    })
                conn.close()
            except:
                pass
    
    return sorted(saves, key=lambda x: x['updated'], reverse=True)


def delete_save(db_path: str) -> bool:
    """删除存档"""
    try:
        if os.path.exists(db_path):
            os.remove(db_path)
            return True
    except:
        pass
    return False
