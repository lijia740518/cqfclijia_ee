from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import sqlite3
import pandas as pd
from io import BytesIO
import os

app = Flask(__name__)
app.secret_key = 'memo_secret_key'  # 用于flash提示信息


# 初始化数据库
def init_db():
    conn = sqlite3.connect('memo.db')
    c = conn.cursor()
    # 创建备忘录表：id(主键)、标题、内容、创建时间
    c.execute('''CREATE TABLE IF NOT EXISTS memos
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  title TEXT NOT NULL,
                  content TEXT,
                  create_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
    conn.commit()
    conn.close()


# 首页：展示所有备忘录
@app.route('/')
def index():
    conn = sqlite3.connect('memo.db')
    conn.row_factory = sqlite3.Row  # 让查询结果可以通过字段名访问
    c = conn.cursor()
    c.execute('SELECT * FROM memos ORDER BY create_time DESC')
    memos = c.fetchall()
    conn.close()
    return render_template('index.html', memos=memos)


# 添加备忘录
@app.route('/add', methods=['POST'])
def add_memo():
    title = request.form['title']
    content = request.form['content']
    if not title:
        flash('标题不能为空！')
        return redirect(url_for('index'))

    conn = sqlite3.connect('memo.db')
    c = conn.cursor()
    c.execute('INSERT INTO memos (title, content) VALUES (?, ?)', (title, content))
    conn.commit()
    conn.close()
    flash('备忘录添加成功！')
    return redirect(url_for('index'))


# 删除备忘录
@app.route('/delete/<int:memo_id>')
def delete_memo(memo_id):
    conn = sqlite3.connect('memo.db')
    c = conn.cursor()
    c.execute('DELETE FROM memos WHERE id = ?', (memo_id,))
    conn.commit()
    conn.close()
    flash('备忘录删除成功！')
    return redirect(url_for('index'))


# 导出备忘录为Excel
@app.route('/export')
def export_memos():
    conn = sqlite3.connect('memo.db')
    # 读取数据到DataFrame
    df = pd.read_sql('SELECT * FROM memos', conn)
    conn.close()

    # 将数据写入BytesIO（内存文件）
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='工作备忘录', index=False)
    output.seek(0)  # 重置文件指针到开头

    # 返回下载响应
    return send_file(
        output,
        download_name='工作备忘录_' + pd.Timestamp.now().strftime('%Y%m%d%H%M%S') + '.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


# 导入Excel备忘录
@app.route('/import', methods=['POST'])
def import_memos():
    if 'file' not in request.files:
        flash('请选择要上传的Excel文件！')
        return redirect(url_for('index'))

    file = request.files['file']
    if file.filename == '':
        flash('文件名称不能为空！')
        return redirect(url_for('index'))

    # 验证文件格式
    if file and file.filename.endswith(('.xlsx', '.xls')):
        try:
            df = pd.read_excel(file)
            # 检查必要列
            if 'title' not in df.columns or 'content' not in df.columns:
                flash('Excel文件必须包含title（标题）和content（内容）列！')
                return redirect(url_for('index'))

            conn = sqlite3.connect('memo.db')
            c = conn.cursor()
            # 批量插入数据
            for _, row in df.iterrows():
                title = row['title']
                content = row['content'] if pd.notna(row['content']) else ''
                if title:  # 标题非空才插入
                    c.execute('INSERT INTO memos (title, content) VALUES (?, ?)', (title, content))
            conn.commit()
            conn.close()
            flash(f'成功导入 {len(df)} 条备忘录（空标题已过滤）！')
        except Exception as e:
            flash(f'导入失败：{str(e)}')
    else:
        flash('仅支持.xlsx/.xls格式的Excel文件！')

    return redirect(url_for('index'))


if __name__ == '__main__':
    init_db()  # 启动时初始化数据库
    app.run(debug=True)  # 调试模式运行，修改代码自动重启