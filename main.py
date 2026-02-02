#модель прогнозирования
import pandas as pd
import numpy as np
import openpyxl
import xlsxwriter
import matplotlib.pyplot as plt
from prophet import Prophet

def prophet(file_exal,column):
    #цикл считывания и обработки всех столбцов файла
    #df_all = pd.read_excel(file_exal)
    df_all = file_exal
    df_all = pd.DataFrame(df_all, columns = ['TRADEDATE',column])
    print(df_all)
    #name_shares = df_all.columns
    #name_shares = name_shares.drop(['TRADEDATE'])
    #print(name_shares)

    #функция для вычисления стандартной метрики SMAPE
    def standard_smape(actual,forecast):
        return round((np.mean(np.abs(forecast - actual) / (np.abs(actual) + np.abs(forecast))))*100,1)

    #количество прогнозных значений
    HORIZONT = 32
    #номер столбца
    i = 1

    df_all['TRADEDATE'] = pd.to_datetime(df_all['TRADEDATE'])
    df_all.columns = ['ds','y']


    #создаём модель Prophet
    model = Prophet()
    #обучаем модель
    model.fit(df_all)

    future = model.make_future_dataframe(periods=32)
    print(future)

    #получаем прогнозы
    forecast = model.predict(future)
    itog = pd.DataFrame(forecast, columns=['ds','yhat'])
    #print(forecast)
    #print(list(forecast))
    print(itog)

    smape = standard_smape(df_all['y'],itog['yhat'][:-32])
    print(f'SMAPE по {i}: {smape:.3f}')

    return forecast, itog




#модуль окна загрузки файла, где выбирается из директории файл, загружается и выводится на экран результат
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import sys
import io

class ExcelViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Predictive Price Model by Arsenteva Anastasiia")
        self.root.geometry("1300x750")
        #self.root.configure(bg="#ffe6f0")
        # --- Визуальный заголовок ---
        title_label = tk.Label(
            root,
            text="📊 tel.8-952-158-2514 готова к сотрудничеству - условия обсуждаемы",
            #bg="#ffe6f0",  # Бледно-розовый фон
            #fg="#800080",  # Тёмно-сиреневый текст
            font=("Arial", 20, "bold"),
            pady=10
        )
        title_label.pack()

        self.df = None
        self.canvas = None
        self.prediction_count = 0

        # --- Кнопки ---
        button_frame = tk.Frame(root, bg="#ffe6f0")
        button_frame.pack(pady=5)

        self.load_button = tk.Button(button_frame, text="📂 Загрузить Excel-файл", command=self.load_excel,  fg="black", font=("Arial", 10, "bold")) #bg="#c8a2c8",
        self.load_button.pack(side=tk.LEFT, padx=5)

        self.refresh_button = tk.Button(button_frame, text="🔄 Обновить", command=self.refresh_table, fg="black", font=("Arial", 10, "bold")) #bg="#c8a2c8",
        self.refresh_button.pack(side=tk.LEFT, padx=5)

        self.predict_button = tk.Button(button_frame, text="🤖 Спрогнозировать", command=self.predict, fg="black", font=("Arial", 10, "bold")) #bg="#c8a2c8",
        self.predict_button.pack(side=tk.LEFT, padx=5)

        self.save_button = tk.Button(button_frame, text="💾 Сохранить в Excel", command=self.save_to_excel, fg="black", font=("Arial", 10, "bold")) #bg="#c8a2c8",
        self.save_button.pack(side=tk.LEFT, padx=5)

        self.exit_button = tk.Button(button_frame, text="❌ Завершить и закрыть", command=self.root.quit, bg="#c8a2c8", fg="black", font=("Arial", 10, "bold"))
        self.exit_button.pack(side=tk.LEFT, padx=5)


        # Диапазон дат для прогноза
        tk.Label(button_frame, text="От:", bg="#ffe6f0").pack(side=tk.LEFT)
        self.date_from_entry = DateEntry(button_frame, width=10, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.date_from_entry.pack(side=tk.LEFT, padx=2)

        tk.Label(button_frame, text="До:", bg="#ffe6f0").pack(side=tk.LEFT)
        self.date_to_entry = DateEntry(button_frame, width=10, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
        self.date_to_entry.pack(side=tk.LEFT, padx=2)

        # Последние N дней для отображения
        tk.Label(button_frame, text="📅 Последние N дней:", bg="#ffe6f0").pack(side=tk.LEFT, padx=(20, 2))
        self.last_n_days_entry = tk.Entry(button_frame, width=5)
        self.last_n_days_entry.insert(0, "0")  # По умолчанию = весь график
        self.last_n_days_entry.pack(side=tk.LEFT)

        # --- Выбор столбца ---
        self.column_selector = ttk.Combobox(button_frame, state="readonly")
        self.column_selector.pack(side=tk.LEFT, padx=10)
        self.column_selector.set("Выберите столбец")

        # --- Таблица ---
        self.table_frame = ttk.Frame(root)
        self.table_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(self.table_frame, show="headings")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        # --- Стилизация заголовков таблицы ---
        style = ttk.Style()
        style.theme_use("default")

        style.configure("Treeview.Heading",
                        background="#ffe6f0",  # Бледно-розовый
                        foreground="black",
                        font=("Arial", 10, "bold"))

        self.x_scroll = tk.Scrollbar(self.table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.y_scroll = tk.Scrollbar(self.table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(xscrollcommand=self.x_scroll.set, yscrollcommand=self.y_scroll.set)
        self.x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.y_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # --- Лог-вывод ---
        log_label = tk.Label(root, text="Лог вывода:")
        log_label.pack()
        self.log_output = tk.Text(root, height=6, bg="black", fg="lime", font=("Courier", 10))
        self.log_output.pack(fill=tk.X, padx=5, pady=5)
        sys.stdout = TextRedirector(self.log_output)

        # --- График ---
        self.plot_frame = tk.Frame(root)
        self.plot_frame.pack(fill=tk.BOTH, expand=True)

    def load_excel(self):
        file_path = filedialog.askopenfilename(
            title="Выберите Excel-файл",
            filetypes=(("Excel файлы", "*.xlsx *.xls"), ("Все файлы", "*.*"))
        )
        if not file_path:
            return

        try:
            self.df = pd.read_excel(file_path)
            print(f"📥 Загружен файл: {file_path}")
            print(f"Столбцы: {list(self.df.columns)}")
            self.column_selector["values"] = list(self.df.columns)
            self.column_selector.set("Выберите столбец")
            self.refresh_table()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось прочитать файл:\n{e}")

    def refresh_table(self):
        if self.df is None:
            return

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(self.df.columns)

        for col in self.df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor=tk.W)

        for _, row in self.df.iterrows():
            values = list(row)
            tags = []
            for col in self.df.columns:
                if col.startswith("Прогноз"):
                    tags.append("highlight")
                    break
            self.tree.insert("", tk.END, values=values, tags=tuple(tags))

        self.tree.tag_configure("highlight", background="#e6ffe6")
        self.update_plot()

    def predict(self):
        if self.df is None:
            messagebox.showwarning("Нет данных", "Сначала загрузите файл.")
            return

        selected_column = self.column_selector.get()
        if selected_column not in self.df.columns:
            messagebox.showwarning("Ошибка", "Выберите столбец для прогноза.")
            return

        try:
            self.prediction_count += 1
            new_col = f"Прогноз {self.prediction_count}"
            diff_col = f"Расхождения_в_% {self.prediction_count}"

            # 🔮 Вызов своей модели Prophet (ожидается [df_full, df_forecast])
            df_predict = prophet(self.df, selected_column)
            df_forecast = df_predict[1]  # df_forecast['ds'], 'yhat'

            # Обработка диапазона дат
            date_from = self.date_from_entry.get()
            date_to = self.date_to_entry.get()

            if date_from:
                df_forecast = df_forecast[df_forecast['ds'] >= pd.to_datetime(date_from)]

            if date_to:
                df_forecast = df_forecast[df_forecast['ds'] <= pd.to_datetime(date_to)]



            # Убедимся, что TRADEDATE — это datetime
            self.df['TRADEDATE'] = pd.to_datetime(self.df['TRADEDATE'])
            df_forecast['ds'] = pd.to_datetime(df_forecast['ds'])

            # Разделим прогноз на:
            existing_dates = set(self.df['TRADEDATE'])
            new_rows = df_forecast[~df_forecast['ds'].isin(existing_dates)].copy()

            # Добавим пустые строки для новых дат
            for _, row in new_rows.iterrows():
                new_entry = {col: None for col in self.df.columns}
                new_entry['TRADEDATE'] = row['ds']
                self.df = pd.concat([self.df, pd.DataFrame([new_entry])], ignore_index=True)

            # Снова сортировка по дате (на всякий случай)
            self.df.sort_values(by='TRADEDATE', inplace=True)
            self.df.reset_index(drop=True, inplace=True)

            # Слияние по дате
            df_forecast = df_forecast[['ds', 'yhat']]
            merged = pd.merge(self.df, df_forecast, how='left', left_on='TRADEDATE', right_on='ds')

            # Добавим столбцы
            self.df[new_col] = merged['yhat']
            self.df[diff_col] = round((1 - self.df[selected_column] / self.df[new_col]) * 100, 1)

            print(f"📊 Добавлен прогноз '{new_col}' с будущими датами.")
            self.refresh_table()

        except Exception as e:
            messagebox.showerror("Ошибка прогноза", str(e))

    def update_plot(self):
        if self.df is None or 'TRADEDATE' not in self.df.columns:
            return

        for widget in self.plot_frame.winfo_children():
            widget.destroy()

        fig, ax = plt.subplots(figsize=(10, 4))

        try:
            # Убедимся в типе даты
            self.df['TRADEDATE'] = pd.to_datetime(self.df['TRADEDATE'])

            # Отображение фактического значения
            selected_column = self.column_selector.get()
            if selected_column in self.df.columns:
                ax.plot(self.df['TRADEDATE'], self.df[selected_column], label=f'Факт: {selected_column}', color='black')

            # Ограничение по последним N дням
            n_days_str = self.last_n_days_entry.get()
            if n_days_str.isdigit() and int(n_days_str) > 0:
                n_days = int(n_days_str)
                cutoff_date = self.df['TRADEDATE'].max() - pd.Timedelta(days=n_days)
                df_plot = self.df[self.df['TRADEDATE'] >= cutoff_date]
            else:
                df_plot = self.df

            # Все прогнозы
            for col in self.df.columns:
                if col.startswith("Прогноз"):
                    ax.plot(df_plot['TRADEDATE'], df_plot[selected_column], label=f'Факт: {selected_column}',
                            color='black')

                    for col in self.df.columns:
                        if col.startswith("Прогноз"):
                            ax.plot(df_plot['TRADEDATE'], df_plot[col], label=col, linestyle='--')

            ax.set_title("📈 Факт vs Прогноз")
            ax.set_xlabel("Дата")
            ax.set_ylabel("Значение")
            ax.grid(True)
            ax.legend()

            fig.autofmt_xdate()

            canvas = FigureCanvasTkAgg(fig, master=self.plot_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        except Exception as e:
            print(f"Ошибка построения графика: {e}")

    def save_to_excel(self):
        if self.df is None:
            messagebox.showwarning("Нет данных", "Сначала загрузите файл.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel файлы", "*.xlsx"), ("Все файлы", "*.*")],
            title="Сохранить файл как..."
        )
        if not file_path:
            return

        try:
            self.df.to_excel(file_path, index=False)
            print(f"✅ Файл сохранён: {file_path}")
            messagebox.showinfo("Готово", "Файл успешно сохранён.")
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", str(e))

class TextRedirector(io.StringIO):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def write(self, s):
        self.text_widget.insert(tk.END, s)
        self.text_widget.see(tk.END)

    def flush(self):
        pass


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelViewerApp(root)
    root.mainloop()