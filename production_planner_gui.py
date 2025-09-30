import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os

class SimpleProductionPlanner:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.load_data()
    
    def load_data(self):
        """Простая загрузка данных"""
        try:
            self.orders_df = pd.read_excel(self.excel_file, sheet_name='Заказы')
            return True
        except Exception as e:
            print(f"Ошибка загрузки: {e}")
            return False

class ProductionPlannerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Production Planner")
        self.root.geometry("800x600")
        
        # Проверка файла
        if not os.path.exists("Объединенная_статистика_заказов.xlsx"):
            messagebox.showerror("Ошибка", "Файл 'Объединенная_статистика_заказов.xlsx' не найден!")
            return
        
        # Загрузка данных
        self.planner = SimpleProductionPlanner("Объединенная_статистика_заказов.xlsx")
        
        # Создание интерфейса
        self.create_ui()
        
    def create_ui(self):
        """Создание простого интерфейса"""
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок
        title_label = ttk.Label(main_frame, text="🏭 Production Planner", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # Информация о данных
        info_label = ttk.Label(main_frame, 
                              text=f"Загружено заказов: {len(self.planner.orders_df)}",
                              font=("Arial", 10))
        info_label.pack(pady=5)
        
        # Таблица заказов
        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Создание Treeview
        columns = ("Номер заказа", "Клиент", "Площадь", "Стоимость")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=15)
        
        # Настройка колонок
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)
        
        # Добавление данных
        for _, row in self.planner.orders_df.iterrows():
            self.tree.insert("", "end", values=(
                row.get("Номер заказа", ""),
                row.get("Клиент", ""),
                row.get("Площадь заказа", 0),
                row.get("Стоимость заказа", 0)
            ))
        
        # Прокрутка
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        
        # Кнопки
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)
        
        ttk.Button(button_frame, text="Рассчитать материалы", 
                  command=self.calculate_materials).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Выход", 
                  command=self.root.quit).pack(side="right", padx=5)
    
    def calculate_materials(self):
        """Простой расчет материалов"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("Информация", "Выберите заказы из таблицы")
            return
        
        order_count = len(selected)
        messagebox.showinfo("Расчет", f"Будет выполнен расчет для {order_count} заказов")

def main():
    root = tk.Tk()
    app = ProductionPlannerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
