"""تشغيل مرة واحدة لتحديث الطالبات؛ يحفظ نسخة احتياطية ولا يمس بنود المخالفات."""
import os, sqlite3, shutil, datetime
from pathlib import Path
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parent
DB=ROOT/'student_violations.db'
XLSX=ROOT/'بيانات_الطالبات_2026-10-22.xlsx'
if not DB.exists():
    raise SystemExit('لم يتم العثور على قاعدة البيانات الأصلية. أوقفي التنفيذ.')
if not XLSX.exists():
    raise SystemExit('لم يتم العثور على ملف الطالبات Excel.')
wb=load_workbook(XLSX,read_only=True,data_only=True)
ws=wb.active
headers=[str(x).strip() if x is not None else '' for x in next(ws.values)]
required=['اسم الطالبة','الصف','رقم هاتف الأب','رقم هاتف الأم']
if any(x not in headers for x in required):
    raise SystemExit('أعمدة ملف Excel غير متطابقة: '+str(headers))
idx={x:i for i,x in enumerate(headers)}
def val(row,k):
    v=row[idx[k]]
    return str(v).strip() if v is not None else ''
def phone(v):
    v=''.join(x for x in v if x.isdigit())
    return ('968'+v) if len(v)==8 else v
students=[];seen=set()
for row in ws.iter_rows(min_row=2,values_only=True):
    name=val(row,'اسم الطالبة');cls=val(row,'الصف')
    if not name:continue
    if not cls:raise SystemExit('طالبة بلا صف: '+name)
    key=(name,cls)
    if key in seen:raise SystemExit('تكرار في ملف Excel: '+name+' '+cls)
    seen.add(key)
    students.append((name,cls,'',phone(val(row,'رقم هاتف الأب')),'',phone(val(row,'رقم هاتف الأم')),''))
if len(students)<100:
    raise SystemExit('عدد الطالبات غير متوقع؛ أُلغي الاستبدال احتياطيًا.')
backup=ROOT/('student_violations_BACKUP_'+datetime.datetime.now().strftime('%Y%m%d_%H%M%S')+'.db')
source=sqlite3.connect(str(DB)); dest=sqlite3.connect(str(backup));source.backup(dest);dest.close();source.close()
conn=sqlite3.connect(str(DB))
try:
    conn.execute('BEGIN IMMEDIATE')
    conn.execute('DELETE FROM violation_records')
    conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='teacher_submissions'")
    if conn.execute("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='teacher_submissions'").fetchone()[0]:
        conn.execute('DELETE FROM teacher_submissions')
    conn.execute('DELETE FROM students')
    conn.executemany('''INSERT INTO students(student_name,class_name,father_name,father_phone,mother_name,mother_phone,area)
       VALUES(?,?,?,?,?,?,?)''',students)
    conn.commit()
except Exception:
    conn.rollback();raise
finally:conn.close()
print(f'تم استبدال القائمة بنجاح: {len(students)} طالبة')
print('النسخة الاحتياطية:',backup)
