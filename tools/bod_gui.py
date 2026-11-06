"""Simple Windows GUI for editing UO BOD Excel + JSON without command line."""
import tkinter as tk
from tkinter import ttk,messagebox
from bod_manager import JOBS,load,save_managed,export_all,next_id


def launch():
 root=tk.Tk();root.title('UO BOD Excel / JSON 資料管理 v3');root.geometry('930x620');root.minsize(800,520)
 frame=ttk.Frame(root,padding=16);frame.pack(fill='both',expand=True)
 top=ttk.Frame(frame);top.pack(fill='x',pady=(0,10))
 ttk.Label(top,text='工匠職業').pack(side='left')
 jobvar=tk.StringVar(value='blacksmith')
 jobsel=ttk.Combobox(top,textvariable=jobvar,values=list(JOBS),state='readonly',width=20);jobsel.pack(side='left',padx=10)
 mode=tk.StringVar(value='rewards')
 ttk.Radiobutton(top,text='獎勵品',variable=mode,value='rewards').pack(side='left',padx=12)
 ttk.Radiobutton(top,text='訂單',variable=mode,value='orders').pack(side='left')
 ttk.Label(frame,text='選擇一筆資料後，可編輯名稱、點數；下方也能新增或刪除。儲存時自動更新 Excel 與 JSON。',foreground='#53677b').pack(anchor='w')
 listframe=ttk.Frame(frame);listframe.pack(fill='both',expand=True,pady=10)
 columns=('id','name','points');tree=ttk.Treeview(listframe,columns=columns,show='headings',selectmode='browse')
 for key,title,width in [('id','ID',90),('name','獎勵名稱／訂單資訊',580),('points','點數',100)]:tree.heading(key,text=title);tree.column(key,width=width,anchor='w' if key=='name' else 'center')
 scroll=ttk.Scrollbar(listframe,orient='vertical',command=tree.yview);tree.configure(yscrollcommand=scroll.set)
 tree.pack(side='left',fill='both',expand=True);scroll.pack(side='left',fill='y')
 form=ttk.Frame(frame);form.pack(fill='x',pady=5)
 fields={}
 for i,(key,label) in enumerate([('name','獎勵名稱／訂單類型'),('material','材料／類型（訂單）'),('qty','每種件數（訂單）'),('quality','品質（訂單）'),('points','點數')]):
  ttk.Label(form,text=label).grid(row=0,column=i,sticky='w',padx=3)
  val=tk.StringVar();e=ttk.Entry(form,textvariable=val,width=28 if i==0 else 14);e.grid(row=1,column=i,sticky='ew',padx=3);fields[key]=val
 form.columnconfigure(0,weight=3)
 status=tk.StringVar(value='');ttk.Label(frame,textvariable=status,foreground='#2563eb').pack(anchor='w',pady=6)
 current={'job':None}
 def refresh():
  try:
   j=load(jobvar.get());current['job']=j
   tree.delete(*tree.get_children())
   for r in j[mode.get()]:
    desc=r['name'] if mode.get()=='rewards' else f"{r['kind']} / {r['material']} / {r['qty']} 件 / {r['quality']}"
    tree.insert('', 'end', iid=r['id'],values=(r['id'],desc,r['points']))
   status.set(f"{j['name']}：{len(j[mode.get()])} 筆。點選一筆即可編輯。")
  except Exception as exc:messagebox.showerror('讀取失敗',str(exc))
 def selected():
  ids=tree.selection()
  if not ids:raise ValueError('請先選擇清單中的資料')
  j=current['job'];return next(v for v in j[mode.get()] if v['id']==ids[0])
 def fill(_=None):
  try:
   r=selected()
   for key,var in fields.items():var.set(r.get('name','') if key=='name' and mode.get()=='rewards' else r.get('kind','') if key=='name' else r.get(key,''))
   if mode.get()=='orders':fields['name'].set(r['kind'])
  except Exception:pass
 def save(add=False,remove=False):
  try:
   j=current['job'];collection=j[mode.get()]
   if remove:
    item=selected()
    if not messagebox.askyesno('確認刪除',f"確定刪除 {item['id']}？"):return
    collection.remove(item)
   else:
    point=int(fields['points'].get())
    if point<0:raise ValueError('點數不可為負數')
    if mode.get()=='rewards':
     name=fields['name'].get().strip()
     if not name:raise ValueError('請輸入獎勵名稱')
     if add:collection.append({'id':next_id(collection,'R',3),'name':name,'points':point,'category':'其他','source':'遊戲實測／手動更新','note':''})
     else:
      r=selected();r.update(name=name,points=point)
    else:
     kind=fields['name'].get().strip();material=fields['material'].get().strip();quality=fields['quality'].get().strip();qty=int(fields['qty'].get())
     if not kind or not material or not quality or qty<=0:raise ValueError('訂單類型、材料、品質及數量都必須有效')
     vals={'kind':kind,'material':material,'quality':quality,'qty':qty,'points':point}
     if add:collection.append({'id':next_id(collection,'O',4),**vals,'source':'遊戲實測／手動更新','note':''})
     else:selected().update(vals)
   save_managed(j);export_all();refresh()
   messagebox.showinfo('更新成功','Excel 與 data/bod_data.json 已同步更新。')
  except Exception as exc:messagebox.showerror('操作失敗',str(exc))
 buttons=ttk.Frame(frame);buttons.pack(fill='x',pady=8)
 ttk.Button(buttons,text='更新選取資料',command=lambda:save()).pack(side='left',padx=4)
 ttk.Button(buttons,text='新增資料',command=lambda:save(add=True)).pack(side='left',padx=4)
 ttk.Button(buttons,text='刪除選取資料',command=lambda:save(remove=True)).pack(side='left',padx=4)
 def just_export():
  try:export_all();messagebox.showinfo('匯出完成','已更新 data/bod_data.json')
  except Exception as e:messagebox.showerror('匯出失敗',str(e))
 ttk.Button(buttons,text='重新匯出 JSON',command=just_export).pack(side='right',padx=4)
 tree.bind('<<TreeviewSelect>>',fill)
 jobsel.bind('<<ComboboxSelected>>',lambda e:refresh())
 mode.trace_add('write',lambda *a:refresh())
 refresh();root.mainloop()
