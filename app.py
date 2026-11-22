# -*- coding: utf-8 -*-
import os, sys, json, datetime, traceback, logging, tempfile, tkinter as tk
from tkinter import ttk, messagebox, filedialog
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

APP='BienBanGiaoNhan'
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
 dlg=tk.Toplevel(root);dlg.title('BienBanGiaoNhan | Đăng nhập');dlg.configure(bg='#f1f5f9');dlg.resizable(False,False)
 dlg.geometry('440x290');dlg.transient(root);dlg.grab_set();dlg.lift();dlg.focus_force()
 shell=tk.Frame(dlg,bg='white',padx=32,pady=24);shell.pack(fill='both',expand=True,padx=18,pady=18)
 tk.Label(shell,text='BIÊN BẢN GIAO NHẬN',bg='white',fg='#0f172a',font=('Segoe UI',17,'bold')).pack(anchor='w')
 tk.Label(shell,text='Đăng nhập để quản lý biểu mẫu 190 / 191',bg='white',fg='#64748b',font=('Segoe UI',10)).pack(anchor='w',pady=(3,19))
 val=tk.StringVar();ent=ttk.Entry(shell,textvariable=val,show='•',font=('Segoe UI',12));ent.pack(fill='x',ipady=7)
 status=tk.Label(shell,text='',bg='white',fg='#dc2626',font=('Segoe UI',9));status.pack(anchor='w',pady=7)
 success=[False]
 def check(*_):
  if val.get()=='024256':success[0]=True;dlg.destroy()
  else:status.configure(text='Mật khẩu không đúng. Vui lòng thử lại.');val.set('');ent.focus_set()
 tk.Button(shell,text='ĐĂNG NHẬP',command=check,bg='#2563eb',fg='white',activebackground='#1d4ed8',activeforeground='white',relief='flat',font=('Segoe UI',11,'bold'),cursor='hand2',pady=9).pack(fill='x',pady=(4,0))
 dlg.bind('<Return>',check);dlg.protocol('WM_DELETE_WINDOW',dlg.destroy);ent.focus_set();root.wait_window(dlg)
 return success[0]

class App:
 def __init__(self,root):
  self.root=root;root.title('BienBanGiaoNhan | Quản lý biên bản');root.geometry('1180x800');root.minsize(880,590)
  self.widgets={};self.current_path=None
  self.style=ttk.Style(root)
  if 'clam' in self.style.theme_names():self.style.theme_use('clam')
  self.style.configure('Modern.TNotebook',background='#f1f5f9',borderwidth=0)
  self.style.configure('Modern.TNotebook.Tab',font=('Segoe UI',10,'bold'),padding=(18,11),background='#e2e8f0',foreground='#334155')
  self.style.map('Modern.TNotebook.Tab',background=[('selected','white')],foreground=[('selected','#1d4ed8')])
  self.style.configure('Modern.TEntry',padding=7,font=('Segoe UI',10))
  root.configure(bg='#f1f5f9')
  header=tk.Frame(root,bg='#0f172a',height=104);header.pack(fill='x');header.pack_propagate(False)
  tk.Label(header,text='BienBanGiaoNhan',bg='#0f172a',fg='white',font=('Segoe UI',23,'bold')).pack(anchor='w',padx=28,pady=(14,0))
  tk.Label(header,text='QUẢN LÝ BIÊN BẢN GIÁM ĐỊNH  •  MẪU 190 & 191',bg='#0f172a',fg='#93c5fd',font=('Segoe UI',10,'bold')).pack(anchor='w',padx=29,pady=(2,0))
  actions=tk.Frame(root,bg='#f1f5f9');actions.pack(fill='x',padx=22,pady=(17,11))
  for label,command,primary in [('＋  Hồ sơ mới',self.clear,False),('▣  Lưu hồ sơ',self.save,False),('▤  Mở hồ sơ',self.load,False),('↓  Xuất mẫu 190',lambda:self.export('190'),True),('↓  Xuất mẫu 191',lambda:self.export('191'),True),('↓  Xuất cả hai',self.export_both,True)]:
   tk.Button(actions,text=label,command=command,bg='#2563eb' if primary else 'white',fg='white' if primary else '#1e293b',activebackground='#1d4ed8' if primary else '#e2e8f0',activeforeground='white' if primary else '#1e293b',relief='flat',bd=0,font=('Segoe UI',10,'bold'),padx=14,pady=11,cursor='hand2').pack(side='left',padx=(0,8))
  body=tk.Frame(root,bg='white',highlightbackground='#e2e8f0',highlightthickness=1);body.pack(fill='both',expand=True,padx=22,pady=(0,12))
  tabs=ttk.Notebook(body,style='Modern.TNotebook');tabs.pack(fill='both',expand=True,padx=15,pady=15)
  for title,fields,subtitle in [('01  THÔNG TIN CHUNG',COMMON,'Thời gian, địa điểm và thông tin kết thúc giao nhận'),('02  BIÊN BẢN MẪU 190',F190,'Mở niêm phong và giao, nhận đối tượng giám định'),('03  BIÊN BẢN MẪU 191',F191,'Đóng gói, niêm phong và giao, nhận lại đối tượng giám định')]:
   page=tk.Frame(tabs,bg='white');tabs.add(page,text=title)
   tk.Label(page,text=subtitle,bg='white',fg='#475569',font=('Segoe UI',11)).pack(anchor='w',padx=20,pady=(17,8))
   container=tk.Frame(page,bg='white');container.pack(fill='both',expand=True)
   canvas=tk.Canvas(container,bg='white',highlightthickness=0);sc=ttk.Scrollbar(container,orient='vertical',command=canvas.yview)
   inner=tk.Frame(canvas,bg='white');win=canvas.create_window((0,0),window=inner,anchor='nw')
   inner.bind('<Configure>',lambda e,c=canvas:c.configure(scrollregion=c.bbox('all')))
   canvas.bind('<Configure>',lambda e,c=canvas,w=win:c.itemconfigure(w,width=e.width))
   canvas.configure(yscrollcommand=sc.set);canvas.pack(side='left',fill='both',expand=True);sc.pack(side='right',fill='y')
   inner.columnconfigure(0,weight=1);inner.columnconfigure(1,weight=1)
   for i,(key,label) in enumerate(fields):
    r=i//2;c=i%2
    cell=tk.Frame(inner,bg='white',padx=18,pady=8);cell.grid(row=r,column=c,sticky='new')
    tk.Label(cell,text=label,bg='white',fg='#334155',font=('Segoe UI',10,'bold'),wraplength=400,justify='left').pack(anchor='w',pady=(0,6))
    if key in MULTI:
     w=tk.Text(cell,height=6 if key.startswith('tinh_trang') else 3,wrap='word',font=('Segoe UI',10),relief='solid',bd=1,highlightthickness=0)
    else:w=ttk.Entry(cell,style='Modern.TEntry',font=('Segoe UI',10))
    w.pack(fill='x');self.widgets[key]=w
  footer=tk.Frame(root,bg='#f1f5f9');footer.pack(fill='x',padx=24,pady=(0,12))
  tk.Label(footer,text='Dữ liệu lưu tại máy của bạn  •  Xuất Word .docx từ biểu mẫu tích hợp',bg='#f1f5f9',fg='#64748b',font=('Segoe UI',9)).pack(side='left')
  tk.Label(footer,text='BienBanGiaoNhan  •  v5.0',bg='#f1f5f9',fg='#64748b',font=('Segoe UI',9)).pack(side='right')
 def collect(self):return {k:(w.get('1.0','end-1c') if isinstance(w,tk.Text) else w.get()) for k,w in self.widgets.items()}
 def populate(self,v):
  for k,w in self.widgets.items():
   if isinstance(w,tk.Text):w.delete('1.0','end');w.insert('1.0',v.get(k,''))
   else:w.delete(0,'end');w.insert(0,v.get(k,''))
 def clear(self):
  if messagebox.askyesno('Hồ sơ mới','Xóa thông tin đang nhập để tạo hồ sơ mới?'):self.populate({});self.current_path=None
 def save(self):
  path=filedialog.asksaveasfilename(defaultextension='.json',filetypes=[('Hồ sơ nhập liệu','*.json')],initialfile='ho_so_giao_nhan.json')
  if path:
   try:
    with open(path,'w',encoding='utf-8') as f:json.dump(self.collect(),f,ensure_ascii=False,indent=2)
    self.current_path=path;messagebox.showinfo('Đã lưu','Đã lưu hồ sơ:\n'+path)
   except Exception as e:messagebox.showerror('Lỗi lưu hồ sơ',str(e))
 def load(self):
  path=filedialog.askopenfilename(filetypes=[('Hồ sơ nhập liệu','*.json')])
  if path:
   try:
    with open(path,encoding='utf-8') as f:v=json.load(f)
    self.populate(v);self.current_path=path
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
  logdir=os.path.join(os.environ.get('LOCALAPPDATA',tempfile.gettempdir()),'BienBanGiamDinh')
  os.makedirs(logdir,exist_ok=True)
  with open(os.path.join(logdir,'startup.log'),'a',encoding='utf-8') as lf:
   lf.write('Starting '+datetime.datetime.now().isoformat()+'\n')
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
  log=os.path.join(os.environ.get('LOCALAPPDATA',tempfile.gettempdir()),'BienBanGiamDinh','BienBanGiamDinh_LOI.txt')
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

