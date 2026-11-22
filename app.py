# -*- coding: utf-8 -*-
import os, sys, json, datetime, traceback, logging, tempfile, tkinter as tk
from tkinter import ttk, messagebox, filedialog
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

APP='Nhập liệu biên bản giám định - Mẫu 190 & 191'
COMMON=[('bat_dau_gio','Giờ bắt đầu'),('bat_dau_phut','Phút bắt đầu'),('bat_dau_ngay','Ngày bắt đầu'),('bat_dau_thang','Tháng bắt đầu'),('bat_dau_nam','Năm bắt đầu'),('dia_diem','Địa điểm lập biên bản'),('ket_thuc_gio','Giờ kết thúc'),('ket_thuc_phut','Phút kết thúc'),('ket_thuc_ngay','Ngày kết thúc'),('ket_thuc_thang','Tháng kết thúc'),('ket_thuc_nam','Năm kết thúc')]
F190=[('qd_so','Số quyết định trưng cầu/giám định bổ sung/giám định lại'),('qd_ngay','Ngày quyết định'),('qd_thang','Tháng quyết định'),('qd_nam','Năm quyết định'),('qd_co_quan','Cơ quan ra quyết định'),('ben_giao','Bên giao (họ tên)'),('ben_giao_co_quan','Bên giao thuộc cơ quan'),('ben_nhan','Bên nhận (họ tên)'),('ben_nhan_co_quan','Bên nhận thuộc cơ quan'),('nguoi_khac','Ông/bà tham gia khác'),('nguoi_khac_chi_tiet','Thông tin người tham gia khác'),('tinh_trang_190','Tình trạng đối tượng gửi giám định (chi tiết)')]
F191=[('nguoi_1','Ông/bà thứ nhất'),('chuc_vu_1','Chức vụ thứ nhất'),('co_quan_1','Cơ quan thứ nhất'),('nguoi_2','Ông/bà thứ hai'),('chuc_vu_2','Chức vụ thứ hai'),('co_quan_2','Cơ quan thứ hai'),('nguoi_3','Ông/bà thứ ba'),('nguoi_3_chi_tiet','Thông tin người thứ ba'),('kl_so','Số kết luận giám định'),('kl_ngay','Ngày kết luận'),('kl_thang','Tháng kết luận'),('kl_nam','Năm kết luận'),('kl_co_quan','Cơ quan ban hành kết luận'),('tinh_trang_191','Tình trạng đối tượng giao nhận lại (chi tiết)')]
MULTI={'tinh_trang_190','tinh_trang_191','nguoi_khac_chi_tiet','nguoi_3_chi_tiet'}
def date(v,p): return f"{v.get(p+'_ngay','')} tháng {v.get(p+'_thang','')} năm {v.get(p+'_nam','')}"
def tm(v,p): return f"{v.get(p+'_gio','')} giờ {v.get(p+'_phut','')}"
def resource_path(name):
 base=getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
 return os.path.join(base,name)

def make_doc(kind,v,path):
 # Load the original converted Word form, retaining its page setup, header, tables and signature layout.
 d=Document(resource_path('template_'+kind+'.docx'))
 ps=d.paragraphs
 def put(idx,txt):
  p=ps[idx]
  # Keep original paragraph formatting and the formatting of the first run.
  # Avoid replacing paragraph objects, which would destroy tab stops, spacing and styles.
  if p.runs:
   p.runs[0].text=txt
   for r in p.runs[1:]: r.text=''
  else: p.add_run(txt)
 def t(key):return str(v.get(key,'')).strip()
 def when(prefix):return f"{t(prefix+'_gio')} giờ {t(prefix+'_phut')} ngày {t(prefix+'_ngay')} tháng {t(prefix+'_thang')} năm {t(prefix+'_nam')}"
 put(6,'Hồi '+when('bat_dau')+' tại: '+t('dia_diem'))
 put(7,'')
 if kind=='190':
  put(8,'Căn cứ Quyết định trưng cầu giám định/giám định bổ sung/giám định lại số: '+t('qd_so')+' ngày '+t('qd_ngay')+' tháng '+t('qd_thang')+' năm '+t('qd_nam')+' của '+t('qd_co_quan'))
  put(10,'Bên giao: '+t('ben_giao'));put(11,'thuộc Cơ quan '+t('ben_giao_co_quan'))
  put(12,'Bên nhận: '+t('ben_nhan'));put(13,'thuộc Cơ quan '+t('ben_nhan_co_quan'))
  put(14,'Ông/bà: '+t('nguoi_khac'));put(15,t('nguoi_khac_chi_tiet'))
  start,end,key=18,26,'tinh_trang_190'
 else:
  put(9,'Ông/bà: '+t('nguoi_1')+'     Chức vụ: '+t('chuc_vu_1'))
  put(10,'thuộc Cơ quan '+t('co_quan_1'))
  put(11,'Ông/bà: '+t('nguoi_2')+'     Chức vụ: '+t('chuc_vu_2'))
  put(12,'thuộc Cơ quan '+t('co_quan_2'))
  put(13,'Ông/bà: '+t('nguoi_3'));put(14,t('nguoi_3_chi_tiet'))
  put(15,'Đã tiến hành đóng gói, niêm phong đối tượng giám định và giao nhận lại đối tượng giám định kèm theo Kết luận giám định số: '+t('kl_so')+' ngày '+t('kl_ngay')+' tháng '+t('kl_thang')+' năm '+t('kl_nam'))
  put(16,'của '+t('kl_co_quan'))
  start,end,key=18,27,'tinh_trang_191'
 lines=t(key).splitlines()
 for i in range(start,end):put(i,lines[i-start] if i-start<len(lines) else '')
 if len(lines)>end-start:
  put(end-1,' '.join(lines[end-start-1:]))
 put(end,'Việc giao, nhận kết thúc hồi '+when('ket_thuc')+'.')
 d.save(path)

def password_gate(root):
 dlg=tk.Toplevel(root);dlg.title('Đăng nhập phần mềm');dlg.resizable(False,False)
 dlg.geometry('370x190');dlg.transient(root);dlg.grab_set()
 ttk.Label(dlg,text='NHẬP MẬT KHẨU SỬ DỤNG',font=('Arial',13,'bold')).pack(pady=(20,12))
 val=tk.StringVar();ent=ttk.Entry(dlg,textvariable=val,show='*',width=30);ent.pack();ent.focus_set()
 status=ttk.Label(dlg,text='',foreground='red');status.pack(pady=6)
 success=[False]
 def check(*_):
  if val.get()=='024256':success[0]=True;dlg.destroy()
  else:status.configure(text='Mật khẩu không đúng');val.set('');ent.focus_set()
 ttk.Button(dlg,text='Đăng nhập',command=check).pack(pady=3)
 dlg.bind('<Return>',check);dlg.protocol('WM_DELETE_WINDOW',dlg.destroy)
 root.wait_window(dlg)
 return success[0]

class App:
 def __init__(self,root):
  self.root=root;root.title(APP);root.geometry('1050x750');self.widgets={};self.values={}
  top=ttk.Frame(root,padding=10);top.pack(fill='x');ttk.Label(top,text=APP,font=('Arial',17,'bold')).pack(anchor='w');ttk.Label(top,text='Lưu hồ sơ JSON • Xuất Word .docx • In bằng Microsoft Word',foreground='#555').pack(anchor='w')
  bar=ttk.Frame(root,padding=(10,0));bar.pack(fill='x')
  for name,fn in [('Hồ sơ mới',self.clear),('Lưu hồ sơ',self.save),('Mở hồ sơ',self.load),('Xuất mẫu 190',lambda:self.export('190')),('Xuất mẫu 191',lambda:self.export('191')),('Xuất cả 2 mẫu',self.export_both)]:ttk.Button(bar,text=name,command=fn).pack(side='left',padx=3,pady=8)
  tabs=ttk.Notebook(root);tabs.pack(fill='both',expand=True,padx=10,pady=8)
  for title,fields in [('Thông tin chung',COMMON),('Mẫu 190',F190),('Mẫu 191',F191)]:
   frame=ttk.Frame(tabs);tabs.add(frame,text=title)
   canvas=tk.Canvas(frame,highlightthickness=0);sc=ttk.Scrollbar(frame,orient='vertical',command=canvas.yview);inner=ttk.Frame(canvas,padding=12);inner.bind('<Configure>',lambda e,c=canvas:c.configure(scrollregion=c.bbox('all')));canvas.create_window((0,0),window=inner,anchor='nw');canvas.configure(yscrollcommand=sc.set);canvas.pack(side='left',fill='both',expand=True);sc.pack(side='right',fill='y');inner.columnconfigure(1,weight=1)
   for row,(key,label) in enumerate(fields):
    ttk.Label(inner,text=label,wraplength=320).grid(row=row,column=0,sticky='nw',padx=(0,15),pady=7)
    if key in MULTI:w=tk.Text(inner,height=7 if key.startswith('tinh_trang') else 3,width=65,wrap='word')
    else:w=ttk.Entry(inner,width=72)
    w.grid(row=row,column=1,sticky='ew',pady=5);self.widgets[key]=w
  ttk.Label(root,text='Xuất Word từ mẫu gốc 190/191; kiểm tra bản in trước khi ký.',padding=8).pack(anchor='w')
 def collect(self):return {k:(w.get('1.0','end-1c') if isinstance(w,tk.Text) else w.get()) for k,w in self.widgets.items()}
 def populate(self,v):
  for k,w in self.widgets.items():
   if isinstance(w,tk.Text):w.delete('1.0','end');w.insert('1.0',v.get(k,''))
   else:w.delete(0,'end');w.insert(0,v.get(k,''))
 def clear(self):
  if messagebox.askyesno('Hồ sơ mới','Xóa thông tin đang nhập để tạo hồ sơ mới?'):self.populate({})
 def save(self):
  path=filedialog.asksaveasfilename(defaultextension='.json',filetypes=[('Hồ sơ nhập liệu','*.json')],initialfile='ho_so_giam_dinh.json')
  if path:
   with open(path,'w',encoding='utf-8') as f:json.dump(self.collect(),f,ensure_ascii=False,indent=2)
   messagebox.showinfo('Đã lưu','Đã lưu hồ sơ: '+path)
 def load(self):
  path=filedialog.askopenfilename(filetypes=[('Hồ sơ nhập liệu','*.json')]);
  if path:
   try:
    with open(path,encoding='utf-8') as f:v=json.load(f)
    self.populate(v)
   except Exception as e:messagebox.showerror('Lỗi mở hồ sơ',str(e))
 def export(self,kind):
  path=filedialog.asksaveasfilename(defaultextension='.docx',filetypes=[('Word','*.docx')],initialfile='Bien_ban_mau_'+kind+'.docx')
  if path:
   try:make_doc(kind,self.collect(),path);messagebox.showinfo('Hoàn tất','Đã xuất Word:\n'+path)
   except Exception as e:messagebox.showerror('Lỗi xuất Word',str(e))
 def export_both(self):
  folder=filedialog.askdirectory(title='Chọn thư mục xuất 2 mẫu')
  if folder:
   try:
    for kind in ['190','191']:make_doc(kind,self.collect(),os.path.join(folder,'Bien_ban_mau_'+kind+'.docx'))
    messagebox.showinfo('Hoàn tất','Đã xuất cả 2 mẫu vào:\n'+folder)
   except Exception as e:messagebox.showerror('Lỗi xuất Word',str(e))
if __name__=='__main__':
 try:
  root=tk.Tk()
  root.withdraw()
  if password_gate(root):
   root.deiconify()
   App(root)
   root.mainloop()
  else:
   root.destroy()
 except Exception:
  import tempfile
  log=os.path.join(os.path.expanduser('~'),'Desktop','BienBanGiamDinh_LOI.txt')
  try:
   with open(log,'w',encoding='utf-8') as f:f.write(traceback.format_exc())
  except Exception:
   log=os.path.join(tempfile.gettempdir(),'BienBanGiamDinh_LOI.txt')
   with open(log,'w',encoding='utf-8') as f:f.write(traceback.format_exc())
  print(traceback.format_exc())
  try:
   from tkinter import messagebox
   messagebox.showerror('Lỗi khởi động','Chi tiết lỗi đã lưu tại: '+log)
  except Exception:pass
  raise

