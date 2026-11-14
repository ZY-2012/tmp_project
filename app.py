"""本地桌面界面：python app.py。密码仅驻留内存。"""
import os
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import date
from aggregate import consolidate

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('集团工资汇总 · 本地工作台')
        self.geometry('940x720')
        self.minsize(820, 660)
        self.events = queue.Queue()
        self.output = None
        self.busy = False
        self.protocol('WM_DELETE_WINDOW', self.close)
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('TFrame', background='#f2f6fa')
        style.configure('TLabel', background='#f2f6fa', foreground='#173856', font=('Microsoft YaHei', 11))
        style.configure('Title.TLabel', font=('Microsoft YaHei', 23, 'bold'))
        style.configure('TButton', font=('Microsoft YaHei', 10), padding=8)
        root = ttk.Frame(self, padding=28)
        root.pack(fill='both', expand=True)
        root.columnconfigure(1, weight=1)
        ttk.Label(root, text='集团工资汇总', style='Title.TLabel').grid(row=0, column=0, columnspan=3, sticky='w')
        ttk.Label(root, text='选择工资表与映射表，一键生成月度付款报告。').grid(row=1, column=0, columnspan=3, sticky='w', pady=(8, 24))
        self.values = {}
        fields = [('month','工资月份',date.today().strftime('%Y-%m')), ('source','工资目录 / 压缩包',''),
                  ('mapping','基础信息映射表',''), ('output','输出Excel',''), ('password','打开密码','')]
        for row, (key, label, default) in enumerate(fields, 2):
            ttk.Label(root, text=label).grid(row=row, column=0, sticky='w', padx=(0, 18), pady=7)
            variable = tk.StringVar(value=default)
            self.values[key] = variable
            ttk.Entry(root, textvariable=variable, show='•' if key == 'password' else '').grid(row=row, column=1, sticky='ew', pady=7)
            if key in ('source','mapping','output'):
                ttk.Button(root, text='浏览', command=lambda k=key: self.browse(k)).grid(row=row, column=2, padx=(12,0))
        choices = ttk.Frame(root)
        choices.grid(row=7, column=0, columnspan=3, sticky='ew', pady=12)
        self.amount = tk.StringVar(value='实发工资')
        ttk.Label(choices, text='金额口径').pack(side='left')
        ttk.Combobox(choices, textvariable=self.amount, values=['实发工资','应发工资'], state='readonly', width=12).pack(side='left', padx=14)
        self.overwrite = tk.BooleanVar()
        self.partial = tk.BooleanVar()
        ttk.Checkbutton(choices, text='允许覆盖成品', variable=self.overwrite).pack(side='left', padx=8)
        ttk.Checkbutton(choices, text='允许不完整报告', variable=self.partial).pack(side='left', padx=8)
        actions = ttk.Frame(root)
        actions.grid(row=8, column=0, columnspan=3, sticky='ew', pady=8)
        self.run_button = ttk.Button(actions, text='开始汇总', command=self.start)
        self.run_button.pack(side='left')
        self.open_button = ttk.Button(actions, text='打开结果目录', command=self.open_output, state='disabled')
        self.open_button.pack(side='left', padx=12)
        self.status = tk.StringVar(value='待开始 · 数据仅在本机处理')
        ttk.Label(root, textvariable=self.status, wraplength=830).grid(row=9, column=0, columnspan=3, sticky='w', pady=12)
        self.bar = ttk.Progressbar(root, mode='determinate')
        self.bar.grid(row=10, column=0, columnspan=3, sticky='ew')
        logframe = ttk.Frame(root)
        logframe.grid(row=11, column=0, columnspan=3, sticky='nsew', pady=14)
        root.rowconfigure(11, weight=1)
        self.log = tk.Text(logframe, height=8, wrap='word', font=('Microsoft YaHei',10), bg='white', relief='flat', state='disabled')
        self.log.pack(side='left', fill='both', expand=True)
        scroll = ttk.Scrollbar(logframe, command=self.log.yview)
        scroll.pack(side='right', fill='y')
        self.log.configure(yscrollcommand=scroll.set)
        ttk.Label(root, text='提示：RAR需要系统解压后端；无法读取时请先手动解压。异常清单与报告保存在同一目录。', wraplength=830).grid(row=12, column=0, columnspan=3, sticky='w')
        self.after(100, self.poll)

    def browse(self, key):
        if key == 'source':
            choice = messagebox.askyesno('选择输入', '选择文件夹？\n点“否”选择 ZIP / RAR 压缩包。')
            path = filedialog.askdirectory() if choice else filedialog.askopenfilename(filetypes=[('压缩包','*.zip *.rar')])
        elif key == 'mapping':
            path = filedialog.askopenfilename(filetypes=[('Excel','*.xls *.xlsx')])
        else:
            path = filedialog.asksaveasfilename(defaultextension='.xlsx', initialfile=f'集团工资汇总_{self.values["month"].get()}.xlsx', filetypes=[('Excel','*.xlsx')])
        if path:
            self.values[key].set(path)

    def append(self, text):
        self.log.configure(state='normal')
        self.log.insert('end', text + '\n')
        self.log.see('end')
        self.log.configure(state='disabled')

    def start(self):
        values = {k:v.get().strip() if k != 'password' else v.get() for k,v in self.values.items()}
        if not all(values[k] for k in ('month','source','mapping','output')):
            messagebox.showerror('请补充', '请填写月份、工资目录、映射表及输出路径。')
            return
        options = dict(password=values['password'] or None, overwrite=self.overwrite.get(), allow_partial=self.partial.get(), amount_field=self.amount.get())
        self.run_button.configure(state='disabled')
        self.open_button.configure(state='disabled')
        self.busy = True
        self.status.set('正在读取工资表…')
        self.append('开始处理。')
        def worker():
            try:
                result, output = consolidate(values['month'], values['source'], values['output'], mapping_path=values['mapping'],
                    progress=lambda i,n,name:self.events.put(('progress', (i,n,name))), **options)
                self.events.put(('done',(result,output)))
            except Exception as error:
                self.events.put(('error',str(error)))
        threading.Thread(target=worker, daemon=True).start()

    def poll(self):
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == 'progress':
                    i,n,name = value
                    self.bar.configure(maximum=n, value=i)
                    self.status.set(f'正在处理 {i}/{n}')
                    self.append(f'{i}/{n}  {name}')
                else:
                    self.busy = False
                    self.values['password'].set('')
                    self.run_button.configure(state='normal')
                    if kind == 'error':
                        self.status.set('处理失败，请查看下方原因')
                        self.append(value)
                        messagebox.showerror('处理失败', value)
                    else:
                        result, self.output = value
                        self.status.set(f'完成 · {len(result.rows)} 行 · 合计 {result.total:,.2f} · 映射 {result.matched}/{len(result.rows)}')
                        self.append(f'报告：{self.output}')
                        for issue in result.issues:
                            self.append(f'{issue["级别"]}：{issue["工作簿"]} / {issue["工作表"]}：{issue["说明"]}')
                        self.open_button.configure(state='normal')
        except queue.Empty:
            pass
        self.after(100, self.poll)

    def close(self):
        if self.busy:
            messagebox.showinfo('正在处理', '请等待当前汇总完成后关闭窗口。')
            return
        self.destroy()

    def open_output(self):
        import subprocess, sys
        folder = str(Path(self.output).parent)
        if os.name == 'nt':
            os.startfile(folder)
        else:
            subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', folder])

if __name__ == '__main__':
    App().mainloop()
