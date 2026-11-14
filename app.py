"""集团工资汇总 v2.0 本地桌面工作台。"""
import os
import sys
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import ttk,filedialog,messagebox
from datetime import date,datetime
from aggregate import consolidate,parse_month,COLUMNS
from preferences import load_preferences,save_preferences
VERSION='2.0'
UI_FONT='Microsoft YaHei UI' if os.name=='nt' else ('PingFang SC' if sys.platform=='darwin' else 'Noto Sans CJK SC')

class App(tk.Tk):
    def __init__(self,preferences_path=None):
        super().__init__()
        self.title('集团工资汇总 v2.0')
        self.geometry(f'{min(1160,self.winfo_screenwidth()-60)}x{max(680,min(850,self.winfo_screenheight()-90))}')
        self.minsize(980,680)
        self.events=queue.Queue()
        self.output=None
        self.busy=False
        self.preference_path=preferences_path
        self.prefs=load_preferences(preferences_path)
        self.protocol('WM_DELETE_WINDOW',self.close)
        style=ttk.Style(self)
        style.theme_use('clam')
        style.configure('.',font=(UI_FONT,10))
        style.configure('TFrame',background='#f1f5f9')
        style.configure('TLabel',background='#f1f5f9',foreground='#16324f')
        style.configure('Title.TLabel',font=(UI_FONT,24,'bold'))
        style.configure('Sub.TLabel',foreground='#61748a')
        style.configure('TButton',padding=(12,7))
        style.configure('Primary.TButton',background='#156a87',foreground='white',padding=(24,10))
        style.map('Primary.TButton',background=[('disabled','#a4b5c0'),('active','#10566e')])
        style.configure('Card.TFrame',background='white')
        style.configure('Card.TLabel',background='white',foreground='#61748a')
        style.configure('Value.TLabel',background='white',foreground='#16324f',font=(UI_FONT,18,'bold'))
        style.configure('Treeview',rowheight=29,background='white',fieldbackground='white')
        style.configure('Treeview.Heading',font=(UI_FONT,10,'bold'),padding=7)
        root=ttk.Frame(self,padding=(24,20))
        root.pack(fill='both',expand=True)
        root.columnconfigure(1,weight=1)
        root.rowconfigure(11,weight=1)
        ttk.Label(root,text='集团工资汇总',style='Title.TLabel').grid(row=0,column=0,columnspan=2,sticky='w')
        ttk.Label(root,text='v2.0 / 本地处理 · 内存解密',style='Sub.TLabel').grid(row=0,column=2,sticky='e')
        ttk.Label(root,text='01 选择来源    →    02 一键汇总    →    03 核对金额与付款信息',style='Sub.TLabel').grid(row=1,column=0,columnspan=3,sticky='w',pady=(6,14))
        modes=ttk.Frame(root)
        modes.grid(row=2,column=0,columnspan=3,sticky='w',pady=(0,8))
        self.mode=tk.StringVar(value=self.prefs.get('mode','month'))
        for text,value in [('按根目录自动定位月份','month'),('直接选择当月目录 / ZIP / RAR','direct')]:
            ttk.Radiobutton(modes,text=text,variable=self.mode,value=value,command=self.update_mode).pack(side='left',padx=(0,24))
        self.source_label=ttk.Label(root,text='工资根目录')
        self.source_label.grid(row=3,column=0,sticky='w',padx=(0,18),pady=5)
        self.source=tk.StringVar()
        ttk.Entry(root,textvariable=self.source).grid(row=3,column=1,sticky='ew')
        ttk.Button(root,text='选择来源…',command=self.browse_source).grid(row=3,column=2,padx=(12,0))
        self.values={}
        for ri,(key,label,hint) in enumerate([('mapping','基础映射表','留空自动查找'),('output','输出Excel','留空保存到当月目录')],4):
            self.values[key]=tk.StringVar(value=self.prefs.get(key,''))
            ttk.Label(root,text=label).grid(row=ri,column=0,sticky='w',pady=5)
            line=ttk.Frame(root)
            line.grid(row=ri,column=1,sticky='ew')
            line.columnconfigure(0,weight=1)
            ttk.Entry(line,textvariable=self.values[key]).grid(row=0,column=0,sticky='ew')
            ttk.Label(line,text=hint,style='Sub.TLabel').grid(row=0,column=1,padx=(10,0))
            ttk.Button(root,text='浏览…',command=lambda k=key:self.browse(k)).grid(row=ri,column=2,padx=(12,0))
        form=ttk.Frame(root)
        form.grid(row=6,column=0,columnspan=3,sticky='ew',pady=(10,6))
        today=date.today()
        months=[f'{y}-{m:02d}' for y in range(today.year+1,today.year-3,-1) for m in range(12,0,-1)]
        self.values['month']=tk.StringVar(value=self.prefs.get('month',today.strftime('%Y-%m')))
        self.amount=tk.StringVar(value=self.prefs.get('amount_field','实发工资'))
        for label,var,options,width,state in [('汇总月份',self.values['month'],months,13,'normal'),('金额口径',self.amount,['实发工资','应发工资'],11,'readonly')]:
            ttk.Label(form,text=label).pack(side='left')
            ttk.Combobox(form,textvariable=var,values=options,width=width,state=state).pack(side='left',padx=(10,22))
        self.values['password']=tk.StringVar()
        ttk.Label(form,text='打开密码').pack(side='left')
        self.password_entry=ttk.Entry(form,textvariable=self.values['password'],show='•',width=18)
        self.password_entry.pack(side='left',padx=10)
        self.show_password=tk.BooleanVar()
        ttk.Checkbutton(form,text='显示',variable=self.show_password,command=lambda:self.password_entry.configure(show='' if self.show_password.get() else '•')).pack(side='left')
        options=ttk.Frame(root)
        options.grid(row=7,column=0,columnspan=3,sticky='w',pady=(2,9))
        self.overwrite=tk.BooleanVar()
        self.partial=tk.BooleanVar()
        self.fallback=tk.BooleanVar(value=True)
        for text,var in [('允许覆盖已有成品',self.overwrite),('Excel占用时自动另存',self.fallback),('允许不完整报告（需人工核对）',self.partial)]:
            ttk.Checkbutton(options,text=text,variable=var).pack(side='left',padx=(0,22))
        actions=ttk.Frame(root)
        actions.grid(row=8,column=0,columnspan=3,sticky='ew',pady=(0,10))
        self.run_button=ttk.Button(actions,text='开始一键汇总',style='Primary.TButton',command=self.start)
        self.run_button.pack(side='left')
        self.open_button=ttk.Button(actions,text='打开汇总表',state='disabled',command=lambda:self.open_output(False))
        self.open_button.pack(side='left',padx=10)
        self.folder_button=ttk.Button(actions,text='打开所在文件夹',state='disabled',command=lambda:self.open_output(True))
        self.folder_button.pack(side='left')
        self.status=tk.StringVar(value='就绪 · 密码不会保存；请确认输入仅含当月数据。')
        ttk.Label(actions,textvariable=self.status,style='Sub.TLabel',wraplength=470).pack(side='right',padx=(18,0))
        self.bar=ttk.Progressbar(root,mode='determinate')
        self.bar.grid(row=9,column=0,columnspan=3,sticky='ew',pady=(0,10))
        cards=ttk.Frame(root)
        cards.grid(row=10,column=0,columnspan=3,sticky='ew',pady=(0,12))
        self.metrics={}
        for i,(key,title) in enumerate([('files','源文件'),('rows','汇总表行数'),('total','工资金额合计'),('mapping','付款信息匹配')]):
            cards.columnconfigure(i,weight=2 if key=='total' else 1)
            card=ttk.Frame(cards,style='Card.TFrame',padding=(14,9))
            card.grid(row=0,column=i,sticky='ew',padx=(0,8 if i<3 else 0))
            ttk.Label(card,text=title,style='Card.TLabel').pack(anchor='w')
            self.metrics[key]=tk.StringVar(value='—')
            ttk.Label(card,textvariable=self.metrics[key],style='Value.TLabel').pack(anchor='w',pady=(3,0))
        tabs=ttk.Notebook(root)
        tabs.grid(row=11,column=0,columnspan=3,sticky='nsew')
        self.preview=self.table(tabs,'汇总预览',COLUMNS,[420,240,140,100,130,150,90])
        self.issues=self.table(tabs,'异常与待补映射',['级别','工作簿','工作表','说明'],[65,400,110,430])
        logframe=ttk.Frame(tabs)
        tabs.add(logframe,text='运行日志')
        self.log=tk.Text(logframe,wrap='word',height=7,bg='#11273a',fg='#e6f0f6',relief='flat',state='disabled',padx=10,pady=10)
        self.log.pack(side='left',fill='both',expand=True)
        scroll=ttk.Scrollbar(logframe,command=self.log.yview)
        scroll.pack(side='right',fill='y')
        self.log.configure(yscrollcommand=scroll.set)
        ttk.Label(root,text='固定七列Excel + 同名 .issues.json 异常清单；新增公司需补齐付款信息后再使用。',style='Sub.TLabel').grid(row=12,column=0,columnspan=3,sticky='w',pady=(10,0))
        self.update_mode()
        self.after(80,self.poll)

    def table(self,tabs,title,columns,widths):
        frame=ttk.Frame(tabs)
        frame.rowconfigure(0,weight=1)
        frame.columnconfigure(0,weight=1)
        tabs.add(frame,text=title)
        tree=ttk.Treeview(frame,columns=columns,show='headings',height=6)
        for name,width in zip(columns,widths):
            tree.heading(name,text=name)
            tree.column(name,width=width,minwidth=65,stretch=False,anchor='e' if name=='金额' else 'w')
        tree.grid(row=0,column=0,sticky='nsew')
        sy=ttk.Scrollbar(frame,command=tree.yview)
        sx=ttk.Scrollbar(frame,orient='horizontal',command=tree.xview)
        sy.grid(row=0,column=1,sticky='ns')
        sx.grid(row=1,column=0,sticky='ew')
        tree.configure(yscrollcommand=sy.set,xscrollcommand=sx.set)
        return tree

    def update_mode(self):
        if hasattr(self, 'last_mode'):
            self.prefs['base_dir' if self.last_mode=='month' else 'source']=self.source.get()
        self.last_mode=self.mode.get()
        key='base_dir' if self.mode.get()=='month' else 'source'
        self.source_label.configure(text='工资根目录' if key=='base_dir' else '当月目录 / 压缩包')
        self.source.set(self.prefs.get(key,''))

    def browse_source(self):
        folder=self.mode.get()=='month' or messagebox.askyesno('选择来源','选择当月文件夹？\n选择“否”可选ZIP或RAR压缩包。')
        path=filedialog.askdirectory() if folder else filedialog.askopenfilename(filetypes=[('压缩包','*.zip *.rar')])
        if path:
            self.source.set(path)
            self.prefs['base_dir' if self.mode.get()=='month' else 'source']=path

    def browse(self,key):
        path=filedialog.askopenfilename(filetypes=[('Excel','*.xls *.xlsx')]) if key=='mapping' else filedialog.asksaveasfilename(defaultextension='.xlsx',initialfile=f'集团工资汇总_{self.values["month"].get()}.xlsx',filetypes=[('Excel','*.xlsx')])
        if path:
            self.values[key].set(path)

    def append(self,text):
        self.log.configure(state='normal')
        self.log.insert('end',f'[{datetime.now():%H:%M:%S}] {text}\n')
        self.log.see('end')
        self.log.configure(state='disabled')

    def start(self):
        if self.busy:
            return
        values={k:(v.get() if k=='password' else v.get().strip()) for k,v in self.values.items()}
        source=self.source.get().strip()
        try:
            parse_month(values['month'])
            if not source or not Path(source).exists():
                raise ValueError('请选择存在的工资根目录、当月目录或压缩包')
        except ValueError as error:
            messagebox.showerror('请检查参数',str(error))
            return
        if self.partial.get() and not messagebox.askyesno('不完整报告','已允许生成不完整报告。该报告需要人工核对，是否继续？'):
            return
        options=dict(password=values['password'] or None,overwrite=self.overwrite.get(),allow_partial=self.partial.get(),amount_field=self.amount.get(),on_locked='timestamp' if self.fallback.get() else 'error')
        options.update(base_dir=source if self.mode.get()=='month' else str(Path(source).parent),input_dir=None if self.mode.get()=='month' else source,mapping_path=values['mapping'] or None,output_path=values['output'] or None)
        self.prefs.update({k:values[k] for k in ('mapping','output','month')})
        self.prefs.update({'base_dir' if self.mode.get()=='month' else 'source':source,'mode':self.mode.get(),'amount_field':self.amount.get()})
        try:
            save_preferences(self.prefs,self.preference_path)
        except OSError:
            self.append('无法保存路径偏好，本次处理不受影响。')
        for button in (self.run_button,self.open_button,self.folder_button):
            button.configure(state='disabled')
        self.busy=True
        self.output=None
        self.preview.delete(*self.preview.get_children())
        self.issues.delete(*self.issues.get_children())
        for metric in self.metrics.values():
            metric.set('—')
        self.bar.configure(value=0)
        self.status.set('正在定位与读取工资表…')
        self.append('开始汇总。')
        def worker():
            try:
                result,output=consolidate(values['month'],**options,progress=lambda i,n,name:self.events.put(('progress',(i,n,name))),log=lambda text:self.events.put(('log',text)))
                self.events.put(('done',(result,output)))
            except Exception as error:
                self.events.put(('error',str(error)))
        threading.Thread(target=worker,daemon=True).start()

    def poll(self):
        try:
            while True:
                kind,value=self.events.get_nowait()
                if kind=='progress':
                    i,n,name=value
                    self.bar.configure(maximum=n,value=i)
                    self.status.set(f'正在处理 {i}/{n} 个源文件…')
                    self.append(f'{i}/{n}  {name}')
                elif kind=='log':
                    self.append(value)
                else:
                    self.busy=False
                    self.values['password'].set('')
                    self.show_password.set(False)
                    self.password_entry.configure(show='•')
                    self.run_button.configure(state='normal')
                    if kind=='error':
                        self.status.set('未生成报告 · 请查看日志与异常清单')
                        self.append(value)
                        messagebox.showerror('汇总失败',value)
                    else:
                        result,self.output=value
                        for key,text in [('files',str(result.files)),('rows',str(len(result.rows))),('total',f'{result.total:,.2f}'),('mapping',f'{result.matched} / {len(result.rows)}')]:
                            self.metrics[key].set(text)
                        for row in result.rows:
                            self.preview.insert('','end',values=[f'{row[c]:,.2f}' if c=='金额' and row[c] is not None else ('' if row[c] is None else row[c]) for c in COLUMNS])
                        for issue in result.issues:
                            self.issues.insert('','end',values=[issue[c] for c in ['级别','工作簿','工作表','说明']])
                        self.append(f'报告：{self.output}\n金额：{result.total:,.2f}；金额缺失：{result.missing_amounts}；待补映射：{len(result.rows)-result.matched}；异常/提示：{len(result.issues)}条。')
                        self.status.set('完成 · 请核对异常与付款信息' if result.issues else '完成 · 全部文件已处理')
                        self.open_button.configure(state='normal')
                        self.folder_button.configure(state='normal')
        except queue.Empty:
            pass
        self.after(80,self.poll)

    def close(self):
        if self.busy:
            messagebox.showinfo('正在汇总','请等待汇总完成后关闭，避免中断文件保存。')
            return
        self.destroy()

    def open_output(self,folder=False):
        import subprocess,sys
        if not self.output or not Path(self.output).is_file():
            messagebox.showerror('文件不存在','汇总表已移动或删除，请重新汇总。')
            return
        path=str(Path(self.output).parent if folder else Path(self.output))
        try:
            if os.name=='nt':
                os.startfile(path)
            else:
                subprocess.Popen(['open' if sys.platform=='darwin' else 'xdg-open',path])
        except OSError as error:
            messagebox.showerror('无法打开',str(error))


def main():
    import sys,tempfile
    if len(sys.argv)==3 and sys.argv[1]=='--self-test':
        with tempfile.TemporaryDirectory() as temp:
            app=App(preferences_path=Path(temp)/'prefs.json')
            app.withdraw()
            app.update()
            assert app.title().endswith('v2.0')
            app.destroy()
        Path(sys.argv[2]).write_text('ok',encoding='utf-8')
        return
    App().mainloop()

if __name__=='__main__':
    main()
