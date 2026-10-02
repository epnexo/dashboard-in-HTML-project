"""Desktop window to clean the CSV and build the reports. Opened by run.bat."""
import os
import queue
import threading
import tkinter as tk
import webbrowser
from tkinter import filedialog, ttk

import pipeline


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Telecom Report')
        self.geometry('820x520')
        self.minsize(600, 420)
        self.queue = queue.Queue()

        frame = ttk.Frame(self, padding=16)
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text='Telecom Transactions Report', font=('Segoe UI', 15, 'bold')).pack(anchor='w')
        ttk.Label(frame, text='Choose the original CSV file, clean it and build the report.').pack(anchor='w', pady=(0, 12))

        row = ttk.Frame(frame)
        row.pack(fill='x')
        self.path = tk.StringVar()
        try:
            self.path.set(str(pipeline.latest_csv()))
        except FileNotFoundError:
            pass
        ttk.Entry(row, textvariable=self.path).pack(side='left', fill='x', expand=True)
        ttk.Button(row, text='Browse…', command=self.browse).pack(side='left', padx=(8, 0))

        self.btn = ttk.Button(frame, text='Clean and build report', command=self.start)
        self.btn.pack(anchor='w', pady=12, ipadx=10, ipady=4)

        self.text = tk.Text(frame, height=12, state='disabled', font=('Consolas', 10), relief='solid', borderwidth=1)
        self.text.pack(fill='both', expand=True)

        bar = ttk.Frame(frame)
        bar.pack(fill='x', pady=(12, 0))
        self.buttons = [
            ttk.Button(bar, text='Open report', command=lambda: webbrowser.open(pipeline.REPORT.as_uri())),
            ttk.Button(bar, text='Open Excel', command=lambda: os.startfile(pipeline.EXCEL)),
            ttk.Button(bar, text='Open cleaning report', command=lambda: webbrowser.open(pipeline.CLEANING_REPORT.as_uri())),
        ]
        for b in self.buttons:
            b.pack(side='left', padx=(0, 8))
        ttk.Button(bar, text='Run history', command=lambda: os.startfile(pipeline.RUNS_DIR)).pack(side='left', padx=(0, 8))
        ttk.Button(bar, text='Open folder', command=lambda: os.startfile(pipeline.ROOT)).pack(side='left')
        pipeline.RUNS_DIR.mkdir(exist_ok=True)
        self.refresh_buttons()
        self.after(100, self.poll)

    def refresh_buttons(self):
        for b, path in zip(self.buttons, (pipeline.REPORT, pipeline.EXCEL, pipeline.CLEANING_REPORT)):
            b.state(['!disabled'] if path.exists() else ['disabled'])

    def browse(self):
        path = filedialog.askopenfilename(initialdir=pipeline.RAW_DIR, filetypes=[('CSV', '*.csv')])
        if path:
            self.path.set(path)

    def write(self, line):
        self.text.configure(state='normal')
        self.text.insert('end', line + '\n')
        self.text.see('end')
        self.text.configure(state='disabled')

    def start(self):
        self.btn.state(['disabled'])
        self.text.configure(state='normal')
        self.text.delete('1.0', 'end')
        self.text.configure(state='disabled')
        threading.Thread(target=self.work, args=(self.path.get().strip() or None,), daemon=True).start()

    def work(self, path):
        # runs off the window thread; it only talks through the queue
        try:
            pipeline.run(path, notify=self.queue.put)
        except Exception:
            pass  # run() already reported the error and logged it in the history
        self.queue.put(None)

    def poll(self):
        try:
            while True:
                line = self.queue.get_nowait()
                if line is None:
                    self.btn.state(['!disabled'])
                    self.refresh_buttons()
                else:
                    self.write(line)
        except queue.Empty:
            pass
        self.after(100, self.poll)


if __name__ == '__main__':
    App().mainloop()
