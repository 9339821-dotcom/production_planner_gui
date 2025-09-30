import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import sys
import json
from collections import defaultdict
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import seaborn as sns

class ModernProductionPlanner:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.orders_df = None
        self.materials_df = None
        self.stock_data = {}
        self.reserved_materials = defaultdict(float)
        self.selected_orders = {}
        self.load_data()
    
    def load_data(self):
        """Загрузка данных из Excel файла"""
        try:
            print("📂 Загрузка данных из Excel файла...")
            
            # Загружаем лист с заказами
            self.orders_df = pd.read_excel(self.excel_file, sheet_name='Заказы')
            
            # Загружаем лист с материалами
            self.materials_df = pd.read_excel(self.excel_file, sheet_name='Потребность материалов')
            
            # Создаем словарь остатков на складе
            if 'На складе' in self.materials_df.columns:
                for _, row in self.materials_df.iterrows():
                    material = row['Материал']
                    if pd.notna(material):
                        stock = row['На складе'] if pd.notna(row['На складе']) else 0
                        self.stock_data[str(material).strip()] = float(stock)
            
            print(f"✅ Загружено: {len(self.orders_df)} заказов, {len(self.materials_df)} материалов")
            
        except Exception as e:
            print(f"❌ Ошибка загрузки данных: {e}")
            raise
    
    def get_companies(self):
        """Получить список компаний"""
        return sorted([str(x) for x in self.orders_df['Клиент'].unique() if pd.notna(x)])

class ModernProductionPlannerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("🏭 Production Planner - Система планирования производства")
        self.root.geometry("1200x800")
        
        # Загрузка данных
        self.planner = None
        self.load_data()
        
        # Создание интерфейса
        self.setup_ui()
    
    def load_data(self):
        """Загрузка данных из файла"""
        excel_file = "Объединенная_статистика_заказов.xlsx"
        if not os.path.exists(excel_file):
            messagebox.showerror("Ошибка", f"Файл {excel_file} не найден!\nПоместите файл в ту же папку, что и программу.")
            self.root.destroy()
            return
        
        try:
            self.planner = ModernProductionPlanner(excel_file)
            messagebox.showinfo("Успех", f"Данные успешно загружены!\nЗаказов: {len(self.planner.orders_df)}\nМатериалов: {len(self.planner.materials_df)}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {e}")
            self.root.destroy()
    
    def setup_ui(self):
        """Создание пользовательского интерфейса"""
        # Главный контейнер
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Заголовок
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = ttk.Label(header_frame, text="🏭 PRODUCTION PLANNER", font=('Arial', 16, 'bold'))
        title_label.pack(pady=10)
        
        # Создание вкладок
        notebook = ttk.Notebook(main_container)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Вкладка 1: Обзор заказов
        self.setup_orders_tab(notebook)
        
        # Вкладка 2: Планирование производства
        self.setup_planning_tab(notebook)
        
        # Вкладка 3: Анализ материалов
        self.setup_materials_tab(notebook)
    
    def setup_orders_tab(self, notebook):
        """Вкладка с обзором заказов"""
        orders_frame = ttk.Frame(notebook)
        notebook.add(orders_frame, text="📋 Заказы")
        
        # Фильтры
        filter_frame = ttk.Frame(orders_frame)
        filter_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(filter_frame, text="Фильтр по компании:", font=('Arial', 10)).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        self.company_var = tk.StringVar()
        companies = ['Все компании'] + self.planner.get_companies()
        company_combo = ttk.Combobox(filter_frame, textvariable=self.company_var, values=companies, state='readonly')
        company_combo.set('Все компании')
        company_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # Таблица заказов
        table_frame = ttk.Frame(orders_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Создание Treeview с прокруткой
        columns = ('Номер', 'Клиент', 'Тип продукции', 'Площадь', 'Стоимость', 'Состояние')
        self.orders_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        # Настройка колонок
        for col in columns:
            self.orders_tree.heading(col, text=col)
            self.orders_tree.column(col, width=120)
        
        # Прокрутка
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.orders_tree.yview)
        self.orders_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.orders_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Загрузка данных
        self.load_orders_data()
        
        # Кнопки действий
        button_frame = ttk.Frame(orders_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(button_frame, text="✅ Выбрать отмеченные для планирования", 
                  command=self.add_selected_orders).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="🔄 Обновить данные", 
                  command=self.load_orders_data).pack(side=tk.LEFT, padx=5)
    
    def setup_planning_tab(self, notebook):
        """Вкладка планирования производства"""
        planning_frame = ttk.Frame(notebook)
        notebook.add(planning_frame, text="📅 Планирование")
        
        # Левая панель - выбранные заказы
        left_frame = ttk.Frame(planning_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        ttk.Label(left_frame, text="Выбранные заказы:", font=('Arial', 12, 'bold')).pack(anchor=tk.W, pady=5)
        
        # Список выбранных заказов
        self.selected_orders_listbox = tk.Listbox(left_frame, height=15, font=('Arial', 10))
        self.selected_orders_listbox.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Кнопки управления выбранными заказами
        order_buttons_frame = ttk.Frame(left_frame)
        order_buttons_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(order_buttons_frame, text="🗑️ Удалить выбранный", 
                  command=self.remove_selected_order).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(order_buttons_frame, text="🧹 Очистить все", 
                  command=self.clear_all_orders).pack(side=tk.LEFT, padx=2)
        
        # Правая панель - управление планированием
        right_frame = ttk.Frame(planning_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Кнопки планирования
        planning_buttons_frame = ttk.Frame(right_frame)
        planning_buttons_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(planning_buttons_frame, text="🧮 Рассчитать потребности", 
                  command=self.calculate_requirements).pack(fill=tk.X, pady=2)
        
        ttk.Button(planning_buttons_frame, text="📦 Зарезервировать материалы", 
                  command=self.reserve_materials).pack(fill=tk.X, pady=2)
        
        ttk.Button(planning_buttons_frame, text="📄 Сформировать заявку на закупку", 
                  command=self.generate_purchase_order).pack(fill=tk.X, pady=2)
        
        # Область результатов
        ttk.Label(right_frame, text="Результаты расчета:", font=('Arial', 12, 'bold')).pack(anchor=tk.W, pady=(20, 5))
        
        self.results_text = scrolledtext.ScrolledText(right_frame, height=10, font=('Arial', 9))
        self.results_text.pack(fill=tk.BOTH, expand=True, pady=5)
    
    def setup_materials_tab(self, notebook):
        """Вкладка анализа материалов"""
        materials_frame = ttk.Frame(notebook)
        notebook.add(materials_frame, text="📦 Материалы")
        
        # Таблица материалов
        table_frame = ttk.Frame(materials_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        columns = ('Материал', 'На складе', 'Зарезервировано', 'Доступно', 'Статус')
        self.materials_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)
        
        for col in columns:
            self.materials_tree.heading(col, text=col)
            self.materials_tree.column(col, width=150)
        
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.materials_tree.yview)
        self.materials_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.materials_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Загрузка данных о материалах
        self.load_materials_data()
    
    def load_orders_data(self):
        """Загрузка данных в таблицу заказов"""
        # Очистка таблицы
        for item in self.orders_tree.get_children():
            self.orders_tree.delete(item)
        
        # Заполнение таблицы
        for _, order in self.planner.orders_df.iterrows():
            self.orders_tree.insert('', tk.END, values=(
                order['Номер заказа'],
                order['Клиент'],
                order.get('Тип продукции', ''),
                f"{order.get('Площадь заказа', 0):.2f}",
                f"{order.get('Стоимость заказа', 0):,.2f}",
                order.get('Состояние заказа', '')
            ))
    
    def add_selected_orders(self):
        """Добавление выбранных заказов в список планирования"""
        selected_items = self.orders_tree.selection()
        if not selected_items:
            messagebox.showwarning("Внимание", "Выберите заказы из таблицы!")
            return
        
        for item in selected_items:
            values = self.orders_tree.item(item)['values']
            order_num = values[0]
            
            # Добавляем в список, если еще нет
            if order_num not in self.get_selected_orders():
                self.selected_orders_listbox.insert(tk.END, order_num)
        
        messagebox.showinfo("Успех", f"Добавлено {len(selected_items)} заказов в планирование")
    
    def get_selected_orders(self):
        """Получить список выбранных заказов"""
        return list(self.selected_orders_listbox.get(0, tk.END))
    
    def remove_selected_order(self):
        """Удаление выбранного заказа из списка"""
        selected_indices = self.selected_orders_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("Внимание", "Выберите заказ для удаления!")
            return
        
        for index in reversed(selected_indices):
            self.selected_orders_listbox.delete(index)
    
    def clear_all_orders(self):
        """Очистка всех выбранных заказов"""
        self.selected_orders_listbox.delete(0, tk.END)
    
    def calculate_requirements(self):
        """Расчет потребностей в материалах"""
        selected_orders = self.get_selected_orders()
        if not selected_orders:
            messagebox.showwarning("Внимание", "Сначала выберите заказы для планирования!")
            return
        
        # Вывод результатов
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, f"📊 РЕЗУЛЬТАТЫ ДЛЯ {len(selected_orders)} ЗАКАЗОВ:\n")
        self.results_text.insert(tk.END, "=" * 50 + "\n\n")
        self.results_text.insert(tk.END, f"Заказы: {', '.join(selected_orders)}\n\n")
        self.results_text.insert(tk.END, "Функция расчета материалов будет реализована в следующей версии.\n")
    
    def reserve_materials(self):
        """Резервирование материалов"""
        selected_orders = self.get_selected_orders()
        if not selected_orders:
            messagebox.showwarning("Внимание", "Сначала выберите заказы!")
            return
        
        messagebox.showinfo("Успех", f"Материалы зарезервированы для {len(selected_orders)} заказов")
    
    def generate_purchase_order(self):
        """Формирование заявки на закупку"""
        selected_orders = self.get_selected_orders()
        if not selected_orders:
            messagebox.showwarning("Внимание", "Сначала выберите заказы!")
            return
        
        purchase_text = f"ЗАЯВКА НА ЗАКУПКУ МАТЕРИАЛОВ\n"
        purchase_text += f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}\n"
        purchase_text += f"Для заказов: {', '.join(selected_orders)}\n\n"
        purchase_text += "Функция формирования заявки будет реализована в следующей версии."
        
        # Показ заявки
        self.show_purchase_order(purchase_text)
    
    def show_purchase_order(self, order_text):
        """Показ заявки на закупку в отдельном окне"""
        order_window = tk.Toplevel(self.root)
        order_window.title("Заявка на закупку")
        order_window.geometry("600x400")
        
        # Текст заявки
        text_widget = scrolledtext.ScrolledText(order_window, font=('Arial', 10), wrap=tk.WORD)
        text_widget.insert(1.0, order_text)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        text_widget.config(state=tk.DISABLED)
        
        # Кнопки
        button_frame = ttk.Frame(order_window)
        button_frame.pack(fill=tk.X, padx=20, pady=10)
        
        ttk.Button(button_frame, text="💾 Сохранить в файл", 
                  command=lambda: self.save_purchase_order(order_text)).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="✖️ Закрыть", 
                  command=order_window.destroy).pack(side=tk.RIGHT, padx=5)
    
    def save_purchase_order(self, order_text):
        """Сохранение заявки в файл"""
        filename = f"Заявка_на_закупку_{datetime.now().strftime('%Y%m%d_%H%M')}.txt"
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(order_text)
            messagebox.showinfo("Успех", f"Заявка сохранена в файл: {filename}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {e}")
    
    def load_materials_data(self):
        """Загрузка данных о материалах"""
        # Очистка таблицы
        for item in self.materials_tree.get_children():
            self.materials_tree.delete(item)
        
        # Заполнение данными
        for material, stock in self.planner.stock_data.items():
            reserved = self.planner.reserved_materials.get(material, 0)
            available = max(0, stock - reserved)
            
            status = "✅ В наличии" if available > 0 else "⚠️ Нет в наличии"
            if reserved > 0:
                status = f"📦 Зарезервировано ({reserved})"
            
            self.materials_tree.insert('', tk.END, values=(
                material,
                f"{stock:.2f}",
                f"{reserved:.2f}",
                f"{available:.2f}",
                status
            ))

def main():
    """Запуск приложения"""
    try:
        root = tk.Tk()
        app = ModernProductionPlannerGUI(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("Критическая ошибка", f"Не удалось запустить приложение: {e}")

if __name__ == "__main__":
    main()
