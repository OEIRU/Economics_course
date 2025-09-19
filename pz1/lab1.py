import tkinter as tk
from tkinter import ttk, messagebox
import csv
from decimal import Decimal, ROUND_HALF_UP

# Функция для округления до 2 знаков
def round2(x):
    return float(Decimal(str(x)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))

class AmortizationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Калькулятор амортизации")
        self.root.geometry("1000x700")
        self.root.resizable(True, True)

        # Переменные
        self.N = tk.StringVar()
        self.tables = {}

        # Интерфейс
        frame_input = tk.Frame(root, padx=10, pady=10)
        frame_input.pack(fill="x")

        tk.Label(frame_input, text="Введите номер в журнале (N):", font=("Arial", 12)).pack(side="left", padx=5)
        tk.Entry(frame_input, textvariable=self.N, width=10, font=("Arial", 12)).pack(side="left", padx=5)
        tk.Button(frame_input, text="Рассчитать", command=self.calculate, bg="#4CAF50", fg="white", font=("Arial", 12)).pack(side="left", padx=10)
        tk.Button(frame_input, text="Сохранить в CSV", command=self.save_to_csv, bg="#2196F3", fg="white", font=("Arial", 12)).pack(side="left", padx=5)

        # Вкладки для таблиц
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Создаём вкладки (пока пустые)
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)
        self.tab3 = ttk.Frame(self.notebook)

        self.notebook.add(self.tab1, text="Линейный способ")
        self.notebook.add(self.tab2, text="Уменьшаемого остатка")
        self.notebook.add(self.tab3, text="Пропорциональный способ")

        # Таблицы внутри вкладок
        self.tree1 = self.create_treeview(self.tab1, ["Год", "Амортизация", "Износ", "Остаточная стоимость"])
        self.tree2 = self.create_treeview(self.tab2, ["Год", "Нач. ост. ст-ть", "Амортизация", "Износ", "Кон. ост. ст-ть"])
        self.tree3 = self.create_treeview(self.tab3, ["Год", "Объём (шт)", "Амортизация", "Износ", "Остаточная стоимость"])

    def create_treeview(self, parent, columns):
        tree = ttk.Treeview(parent, columns=columns, show="headings", height=10)
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, anchor="center", width=120)
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        return tree

    def calculate(self):
        try:
            N = int(self.N.get())
            if N <= 0:
                raise ValueError
        except:
            messagebox.showerror("Ошибка", "Введите корректный положительный номер N!")
            return

        cost = N * 1000.0
        years = 5
        volumes = [150, 350, 600, 300, 200]  # по годам

        # 1. Линейный способ
        amort_linear = cost / years
        table1 = []
        accumulated = 0.0
        for year in range(1, years+1):
            accumulated += amort_linear
            residual = cost - accumulated
            table1.append([year, round2(amort_linear), round2(accumulated), round2(residual)])

        # 2. Уменьшаемого остатка (k=1.25)
        k = 1.25
        rate = (1 / years) * k
        table2 = []
        residual_start = cost
        accumulated = 0.0
        for year in range(1, years+1):
            if year == years:
                amort = residual_start  # списываем остаток
            else:
                amort = residual_start * rate
            accumulated += amort
            residual_end = residual_start - amort
            table2.append([year, round2(residual_start), round2(amort), round2(accumulated), round2(residual_end)])
            residual_start = residual_end

        # 3. Пропорциональный способ
        total_volume = sum(volumes)
        amort_per_unit = cost / total_volume
        table3 = []
        accumulated = 0.0
        for year in range(1, years+1):
            vol = volumes[year-1]
            amort = vol * amort_per_unit
            accumulated += amort
            residual = cost - accumulated
            table3.append([year, vol, round2(amort), round2(accumulated), round2(residual)])

        # Сохраняем для экспорта
        self.tables = {
            "linear": table1,
            "declining": table2,
            "units": table3,
            "cost": cost
        }

        # Очищаем таблицы
        for tree in [self.tree1, self.tree2, self.tree3]:
            for item in tree.get_children():
                tree.delete(item)

        # Заполняем таблицы
        for row in table1:
            self.tree1.insert("", "end", values=row)

        for row in table2:
            self.tree2.insert("", "end", values=row)

        for row in table3:
            self.tree3.insert("", "end", values=row)

        messagebox.showinfo("Успех", f"Расчёты для N={N} (стоимость = {cost} у.е.) успешно выполнены!")

    def save_to_csv(self):
        if not hasattr(self, 'tables') or not self.tables:
            messagebox.showwarning("Предупреждение", "Сначала выполните расчёты!")
            return

        cost = self.tables["cost"]
        N = int(self.N.get())

        # 1. Линейный
        with open(f"amortization_linear_N{N}.csv", "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Год", "Амортизация за год", "Накопленный износ", "Остаточная стоимость"])
            for row in self.tables["linear"]:
                writer.writerow(row)

        # 2. Уменьшаемого остатка
        with open(f"amortization_declining_N{N}.csv", "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Год", "Ост. ст-ть на начало", "Амортизация", "Накопленный износ", "Ост. ст-ть на конец"])
            for row in self.tables["declining"]:
                writer.writerow(row)

        # 3. Пропорциональный
        with open(f"amortization_units_N{N}.csv", "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Год", "Объём продукции (шт)", "Амортизация", "Накопленный износ", "Остаточная стоимость"])
            for row in self.tables["units"]:
                writer.writerow(row)

        messagebox.showinfo("Экспорт", f"Таблицы успешно сохранены в файлы:\n"
                                       f"amortization_linear_N{N}.csv\n"
                                       f"amortization_declining_N{N}.csv\n"
                                       f"amortization_units_N{N}.csv")

# Запуск приложения
if __name__ == "__main__":
    root = tk.Tk()
    app = AmortizationApp(root)
    root.mainloop()