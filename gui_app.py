"""
Графический интерфейс для статистического анализа.
Упаковывается в EXE через PyInstaller.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import threading

def resource_path(relative_path):
    """Возвращает абсолютный путь к ресурсу, работает и в exe, и в скрипте."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

class StatAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title(" Статистический анализ: До vs После")
        self.root.geometry("750x600")
        self.root.resizable(False, False)

        # Стиль
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Title.TLabel', font=('Segoe UI', 16, 'bold'), foreground='white')
        style.configure('TButton', font=('Segoe UI', 11), padding=8)
        style.configure('Header.TFrame', background='#3498db')

        self._build_ui()

    def _build_ui(self):
        # Заголовок
        header = ttk.Frame(self.root, style='Header.TFrame')
        header.pack(fill='x')
        ttk.Label(header, text=" Статистический анализ: До vs После",
                  style='Title.TLabel', background='#3498db', foreground='white').pack(pady=15)

        main = ttk.Frame(self.root, padding=20)
        main.pack(fill='both', expand=True)

        # Шаг 1
        ttk.Label(main, text="Шаг 1. Создайте шаблон файла и заполните его данными",
                  font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(0, 5))
        ttk.Button(main, text=" Создать шаблон Excel", command=self.create_template).pack(fill='x', pady=(0, 20))

        # Шаг 2
        ttk.Label(main, text="Шаг 2. Выберите заполненный файл", font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(0, 5))
        file_frame = ttk.Frame(main)
        file_frame.pack(fill='x', pady=(0, 5))
        self.file_var = tk.StringVar(value="(файл не выбран)")
        ttk.Entry(file_frame, textvariable=self.file_var, state='readonly').pack(side='left', fill='x', expand=True)
        ttk.Button(file_frame, text="📂 Выбрать", width=12, command=self.choose_file).pack(side='right', padx=(5, 0))

        # Шаг 3 (По умолчанию папка Загрузки)
        default_download_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        if not os.path.exists(default_download_dir):
            default_download_dir = os.path.expanduser("~")
            
        ttk.Label(main, text="Шаг 3. Папка для отчёта и графиков", font=('Segoe UI', 10, 'bold')).pack(anchor='w', pady=(10, 5))
        out_frame = ttk.Frame(main)
        out_frame.pack(fill='x', pady=(0, 20))
        self.out_var = tk.StringVar(value=default_download_dir)
        ttk.Entry(out_frame, textvariable=self.out_var, state='readonly').pack(side='left', fill='x', expand=True)
        ttk.Button(out_frame, text="📁 Изменить", width=12, command=self.choose_output).pack(side='right', padx=(5, 0))

        # Кнопка запуска
        self.run_btn = ttk.Button(main, text=" ЗАПУСТИТЬ АНАЛИЗ", command=self.run_analysis)
        self.run_btn.pack(fill='x', pady=10, ipady=5)

        # Прогресс и лог
        self.progress = ttk.Progressbar(main, mode='indeterminate')
        self.progress.pack(fill='x', pady=(10, 5))
        self.log = tk.Text(main, height=8, font=('Consolas', 9), bg='#f8f9fa', relief='sunken')
        self.log.pack(fill='both', expand=True)

    def log_msg(self, text):
        self.log.insert('end', text + '\n')
        self.log.see('end')
        self.root.update_idletasks()

    def create_template(self):
        path = filedialog.asksaveasfilename(title="Сохранить шаблон", defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")], initialfile="шаблон_данных.xlsx")
        if not path: return
        try:
            import stat_analysis
            stat_analysis.generate_template(path)
            self.log_msg(f"✅ Шаблон создан: {path}")
            messagebox.showinfo("Успех", f"Шаблон создан:\n{path}\n\nЗаполните его своими данными.")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать шаблон:\n{e}")

    def choose_file(self):
        path = filedialog.askopenfilename(title="Выберите файл данных", filetypes=[("Excel", "*.xlsx *.xls")])
        if path: self.file_var.set(path)

    def choose_output(self):
        path = filedialog.askdirectory(title="Папка для результатов", initialdir=self.out_var.get())
        if path: self.out_var.set(path)

    def run_analysis(self):
        file_path = self.file_var.get()
        out_dir = self.out_var.get()
        if not os.path.isfile(file_path):
            messagebox.showerror("Ошибка", "Выберите файл данных!")
            return
        
        self.run_btn.config(state='disabled')
        self.progress.start()
        self.log.delete('1.0', 'end')
        self.log_msg("▶ Запуск анализа...")

        thread = threading.Thread(target=self._analysis_worker, args=(file_path, out_dir), daemon=True)
        thread.start()

    def _analysis_worker(self, file_path, out_dir):
        try:
            import stat_analysis
            if not hasattr(stat_analysis, 'run_analysis'):
                raise ImportError("В stat_analysis.py нет функции run_analysis")
            
            results = stat_analysis.run_analysis(file_path, out_dir, log_callback=self.log_msg)
            
            self.root.after(0, lambda: self.log_msg(f"\n✅ Анализ завершён!\n📄 Отчёт: {results['report_path']}\n📊 Графики: {out_dir}"))
            self.root.after(0, lambda: messagebox.showinfo("Готово", f"Анализ завершён!\n\nОтчёт: {results['report_path']}\nГрафики: {out_dir}"))
        except Exception as e:
            import traceback
            err = traceback.format_exc()
            self.root.after(0, lambda: self.log_msg(f"\n❌ Ошибка: {e}\n{err}"))
            self.root.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            self.root.after(0, self.progress.stop)
            self.root.after(0, lambda: self.run_btn.config(state='normal'))

if __name__ == "__main__":
    root = tk.Tk()
    app = StatAnalysisApp(root)
    root.mainloop()