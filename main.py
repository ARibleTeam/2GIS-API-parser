import math
from void import *
import tkinter as tk
from tkinter import ttk
import os
import subprocess

# Использование глобальных переменных - ОЧЕНЬ ПЛОХО в контексте данной программы. 
# Однако, проект используется для изучения ДРУГИМИ студентами - рефакторинг часть их задачи

KEY = None
MAX_ORG_LIST = 5

city_input = False

city = ""
selected_region = ""
parent_ids = []
all_category_name_id = {}
# ключ-значение. Имя - айди
all_category_id_orgs = {}
# ключ-значение. айди - количество организаций

selected_categories = []
total_org_count = 0

organizations = {}
# ключ-значение. айди организации - данные организации
organization_products = {}
# Словарь для хранения данных о продуктах по организациям


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("2GIS Magic Parser")
        self.geometry("960x540")

        # Контейнер для экранов
        container = tk.Frame(self)
        container.pack(fill="both", expand=True)

        # Словарь для хранения экранов
        self.frames = {}
        self.container = container  # Сохраняем контейнер для дальнейшего использования

        self.show_frame(StartPage)

    def show_frame(self, page_class):
        # Создаем экран только если его еще нет в self.frames
        if page_class not in self.frames:
            frame = page_class(parent=self.container, controller=self)
            self.frames[page_class] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        # Поднимаем текущий экран
        frame = self.frames[page_class]
        frame.tkraise()


class StartPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        self.controller = controller
        self.buttons = []  # Список для кнопок, чтобы скрывать их после обновления
        self.error_label = None  # Для хранения метки с ошибкой

        # Проверка наличия файла key.txt и чтение ключа
        self.check_api_key()

        # Центрирование содержимого
        self.pack_propagate(False)
        self.config(width=960, height=540)

        if KEY:
            # Если ключ найден, показываем форму для ввода города
            self.show_city_form()
        else:
            # Если ключ пуст или не найден, показываем сообщение
            self.show_key_error()

    def check_api_key(self):
        """Проверка наличия и чтения ключа API из файла"""
        global KEY
        if os.path.exists("key.txt"):
            with open("key.txt", "r") as f:
                KEY = f.readline().strip()
        if not KEY:
            KEY = None  # Если ключ пустой, то он None

    def show_city_form(self):
        """Показываем форму для ввода города"""
        label = tk.Label(self, text="Введите название города", font=("Arial", 20, "bold"))
        label.pack(pady=(80, 20))

        # Поле для ввода города
        self.entry = tk.Entry(self, width=40, font=("Arial", 14))
        self.entry.pack(pady=10)

        # Кнопка "Далее"
        def on_button_click():
            global city
            city = self.entry.get().strip()
            if city:
                self.controller.show_frame(PageOne)
            else:
                self.show_error("Пожалуйста, введите название города.")

        button1 = ttk.Button(self, text="Далее", command=on_button_click)
        button1.pack(pady=(30, 0))

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Размещаем содержимое на центральной позиции в окне
        label.place(relx=0.5, rely=0.3, anchor="center")
        self.entry.place(relx=0.5, rely=0.45, anchor="center")
        button1.place(relx=0.5, rely=0.6, anchor="center")

    def show_key_error(self):
        """Показываем ошибку и кнопки для работы с ключом"""
        self.error_label = tk.Label(self, text="Ошибка: Не найден ключ API! Пожалуйста, введите ключ в файл key.txt.",
                                    fg="red", font=("Arial", 14, "bold"))
        self.error_label.pack(pady=50)

        # Кнопки для работы с файлом key.txt
        def update_key():
            self.check_api_key()
            # Если ключ найден, скрываем кнопки и показываем форму
            if KEY:
                self.error_label.pack_forget()  # Скрываем метку с ошибкой
                for button in self.buttons:  # Скрываем все кнопки
                    button.pack_forget()
                self.show_city_form()

        def open_file():
            if not os.path.exists("key.txt"):
                with open("key.txt", "w") as f:
                    f.write("")  # Создаем пустой файл, если его нет
            subprocess.run(["notepad", "key.txt"])  # Открытие файла в блокноте

        def close_program():
            self.controller.quit()

        button_update = ttk.Button(self, text="Обновить", command=update_key)
        button_update.pack(pady=5)
        self.buttons.append(button_update)

        button_open = ttk.Button(self, text="Открыть файл", command=open_file)
        button_open.pack(pady=5)
        self.buttons.append(button_open)

        button_close = ttk.Button(self, text="Завершить", command=close_program)
        button_close.pack(pady=5)
        self.buttons.append(button_close)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=1)

    def show_error(self, message):
        error_label = tk.Label(self, text=message, fg="red", font=("Arial", 12, "italic"))
        error_label.place(relx=0.5, rely=0.55, anchor="center")

class PageOne(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        # Устанавливаем размер окна и центрируем элементы
        self.pack_propagate(False)
        self.config(width=960, height=540)

        self.regions = {}

        # Заголовок
        label = tk.Label(self, text="Выберите регион из предложенных:", font=("Arial", 20, "bold"))
        label.pack(pady=(40, 20))  # Отступ сверху для более удобного размещения текста

        # Создаем Frame для Listbox и Scrollbar
        list_frame = tk.Frame(self)
        list_frame.pack(pady=(10, 30), fill=tk.BOTH, expand=True)

        # Создаем Listbox и подключаем к нему Scrollbar
        self.listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE, font=("Arial", 14), width=50, height=12)
        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        self.listbox.config(yscrollcommand=scrollbar.set)

        # Расположение Listbox и Scrollbar
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(20, 0))  # Отступ слева для Listbox
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 20))  # Отступ справа для Scrollbar

        # Заполняем Listbox регионами
        global city
        if len(city) > 3:
            for item in getregionsids(city, KEY)["result"]["items"]:
                region_name = item["name"]
                region_id = item["id"]
                # Заполняем словарь
                self.regions[region_name] = region_id
                # Добавляем название региона в Listbox
                self.listbox.insert(tk.END, region_name)

        # Кнопка для подтверждения выбора региона
        button = ttk.Button(self, text="Далее", command=lambda: self.select_region(controller))
        button.pack(pady=(20, 40))  # Отступ перед кнопкой и снизу

        # Центрируем элементы по вертикали и горизонтали
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Размещаем элементы по центру окна
        label.place(relx=0.5, rely=0.2, anchor="center")
        list_frame.place(relx=0.5, rely=0.5, anchor="center")
        button.place(relx=0.5, rely=0.85, anchor="center")

    def select_region(self, controller):
        # Получаем индекс выбранного элемента
        selected_index = self.listbox.curselection()
        if selected_index:
            # Получаем название выбранного региона
            selected_region_name = self.listbox.get(selected_index)
            # Получаем ID региона из словаря
            selected_region_id = self.regions[selected_region_name]
            # Сохраняем ID в глобальную переменную или используем по необходимости
            global selected_region
            selected_region = selected_region_id
            controller.show_frame(PageTwo)


class PageTwo(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        # Заголовок
        label = tk.Label(self, text="Выберите категории для парсинга", font=("Arial", 16))
        label.pack(pady=20)

        # Глобальная переменная для выбранного региона
        global selected_region

        if len(selected_region) >= 1:
            # Извлекаем данные категорий
            for item in get_parent_categories(selected_region, KEY)["result"]["items"]:
                parent_ids.append(item["id"])


            for parent_id in parent_ids:
                data = get_categories_by_parent_id(selected_region, parent_id, KEY)
                for item in data["result"]["items"]:
                    all_category_name_id[item["name"]] = item["id"]
                    all_category_id_orgs[item["id"]] = item.get("branch_count", "0")


        # Frame для поиска и Listbox
        search_frame = tk.Frame(self)
        search_frame.pack(pady=10, fill=tk.X)

        # Поле для ввода текста (поиск)
        self.search_entry = tk.Entry(search_frame, font=("Arial", 12))
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, padx=10, expand=True)

        # Frame для Listbox и Scrollbar
        list_frame = tk.Frame(self)
        list_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        # Listbox с прокруткой
        self.listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE, font=("Arial", 12), width=100, height=10)
        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        self.listbox.config(yscrollcommand=scrollbar.set)

        # Заполняем Listbox названиями категорий
        self.update_listbox(all_category_name_id)

        # Размещение Listbox и Scrollbar
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Метка для отображения общего количества организаций
        self.total_label = tk.Label(self, text="Кол-во организаций для парсинга: 0", font=("Arial", 14))
        self.total_label.pack(pady=10)

        # Кнопка для подтверждения выбора
        select_button = ttk.Button(self, text="Далее", command=lambda: self.click(controller))
        select_button.pack(pady=10)

        # Привязка события выбора элемента к обновлению количества организаций
        self.listbox.bind('<ButtonRelease-1>', self.on_item_click)

        # Привязка события ввода в поле поиска
        self.search_entry.bind("<KeyRelease>", self.filter_categories)

    def click(self, controller):
        controller.show_frame(PageThree)

    def on_item_click(self, event):
        # Получаем индекс выбранного элемента, по которому кликнули
        index = self.listbox.nearest(event.y)  # Получаем индекс элемента по вертикальной координате клика
        category_text = self.listbox.get(index)
        category_name = category_text.split(" (")[0]
        category_id = all_category_name_id[category_name]
        org_count = int(all_category_id_orgs[category_id])

        # Проверяем, если категория уже в selected_categories
        if any(category["id"] == category_id for category in selected_categories):
            # Удаляем категорию из selected_categories
            selected_categories[:] = [category for category in selected_categories if category["id"] != category_id]
        else:
            # Добавляем категорию в selected_categories
            selected_categories.append({
                "id": category_id,
                "name": category_name,
                "organizations": org_count
            })

        # Обновляем Listbox после изменения выбранных категорий
        self.update_listbox(all_category_name_id)

        # Обновляем общее количество организаций
        self.update_total()

    def update_total(self):
        global total_org_count
        total_org_count = 0  # Сбросить перед подсчетом

        # Перебираем выбранные категории
        for category in selected_categories:
            total_org_count += category["organizations"]

        print(selected_categories)

        # Обновляем метку с общим количеством организаций
        self.total_label.config(text=f"Кол-во организаций для парсинга: {total_org_count}")

    def filter_categories(self, event=None):
        search_text = self.search_entry.get().lower()

        # Фильтруем категории по введенному тексту
        filtered_categories = {name: category_id for name, category_id in all_category_name_id.items()
                               if search_text in name.lower()}

        # Обновляем Listbox с отфильтрованными категориями
        self.update_listbox(filtered_categories)

    def update_listbox(self, categories):
        # Очистить текущий список в Listbox
        self.listbox.delete(0, tk.END)

        # Заполняем Listbox новыми категориями
        for category_name, category_id in categories.items():
            org_count = all_category_id_orgs.get(category_id, "0")
            display_text = f"{category_name} (ID: {category_id}, Орг.: {org_count})"
            self.listbox.insert(tk.END, display_text)

            # Проверяем, если категория уже выбрана, выделяем её красным фоном
            index = self.listbox.size() - 1  # Получаем индекс только что добавленного элемента
            if any(category["id"] == category_id for category in selected_categories):
                self.listbox.itemconfig(index, {'bg': 'red'})
            else:
                self.listbox.itemconfig(index, {'bg': 'white'})


class PageThree(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        # Проверяем, выбраны ли категории
        if len(selected_categories) > 0:
            # Заголовок
            label = tk.Label(self, text="Сбор данных...", font=("Arial", 20, "bold"))
            label.pack(pady=(50, 20))  # Увеличенный отступ сверху

            # Метка прогресса для категорий
            self.progress_label_cats = tk.Label(self, text="0/0 категорий...", font=("Arial", 14))
            self.progress_label_cats.pack(pady=10)

            # Метка прогресса для организаций
            self.progress_label_orgs = tk.Label(self, text="0/0 организаций...", font=("Arial", 14))
            self.progress_label_orgs.pack(pady=10)

            # Переменные для отслеживания прогресса
            count = 0
            count_cat = 0
            count_orgs_req = 0
            count_products_req = 0

            def process_data():
                nonlocal count, count_orgs_req, count_products_req, count_cat


                for category in selected_categories:
                    count_orgs_req = 0
                    count_cat += 1
                    self.progress_label_cats.config(text=f"{count_cat}/{len(selected_categories)} категорий...")
                    self.update()
                    page = 1
                    category_id = category["id"]
                    org_count = category["organizations"]
                    org_per_page = 10  # Количество организаций на одной странице

                    pages_required = math.ceil(org_count / org_per_page)

                    while page <= MAX_ORG_LIST and page <= pages_required:
                        count += 1
                        # Обновляем метки прогресса
                        data = get_org_by_rubric_id(category_id, selected_region, page, KEY)

                        items = data.get("result", {}).get("items", [])
                        total = data.get("result", {}).get("total", 0)
                        for item in items:
                            count_orgs_req += 1
                            self.progress_label_orgs.config(text=f"{count_orgs_req}/{org_count} организаций...")
                            self.update()
                            org_id = item.get("id", "Отсутствует")
                            contacts_data = getcontacts(org_id)

                            organizations[org_id] = {
                                "name": item.get("name", "Отсутствует"),
                                "address_name": item.get("address_name", "Отсутствует"),
                                "address_comment": item.get("address_comment", "Отсутствует"),
                                "category_id": category_id,
                                "phones": contacts_data.get("phones", ["-"]),
                                "emails": contacts_data.get("emails", ["-"]),
                            }

                        if count % 400 == 0 or count_orgs_req % 400 == 0:
                            time.sleep(20)

                        page += 1

                # Обработка продуктов для организаций
                for organization_id in organizations.keys():
                    count_products_req += 1
                    data = getproducts(organization_id)
                    products = []
                    for item in data["result"]["items"]:
                        product = item["product"]
                        offer = item["offer"]

                        product_info = {
                            "name": product.get("name", "Отсутствует"),
                            "description": product.get("description", "Отсутствует"),
                            "price": offer.get("price", "Отсутствует"),
                        }
                        products.append(product_info)

                    organization_products[organization_id] = products

                    if count_products_req % 400 == 0:
                        time.sleep(20)

                # Завершаем обработку
                label.config(text="Данные успешно собраны!", font=("Arial", 18, "bold"), fg="green")
                self.progress_label_cats.config(text="Все категории собраны", fg="green")
                self.progress_label_orgs.config(text="Все организации собраны", fg="green")
                self.update()

                time.sleep(2)  # Короткая пауза перед переходом
                controller.show_frame(PageFour)  # Переход на следующий экран

            # Запуск обработки данных в фоновом потоке
            self.after(100, process_data)


class PageFour(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)

        # Вызываем функцию создания файла Excel
        createXLSX(organization_products, organizations, selected_categories)
        print("Таблица Excel создана и сохранена как 'organizations.xlsx'.")

        # Заголовок
        label = tk.Label(self, text="Данные успешно собраны и сохранены в файл Excel!", font=("Arial", 18, "bold"), fg="green")
        label.pack(pady=(50, 20))  # Отступ сверху

        # Описание файла
        info_label = tk.Label(self, text="Файл сохранен как 'organizations.xlsx' в текущей директории.", font=("Arial", 14))
        info_label.pack(pady=10)

        # Кнопка для открытия Excel-файла
        open_button = ttk.Button(self, text="Открыть файл Excel", command=self.open_excel_file)
        open_button.pack(pady=10)

        # Кнопка для завершения работы
        finish_button = ttk.Button(self, text="Завершить", command=lambda: controller.quit())
        finish_button.pack(pady=10)

    def open_excel_file(self):
        # Открываем созданный Excel файл
        try:
            if os.name == 'nt':  # Windows
                os.startfile("organizations.xlsx")
            elif os.name == 'posix':  # macOS или Linux
                subprocess.call(["open" if os.uname().sysname == 'Darwin' else "xdg-open", "organizations.xlsx"])
        except Exception as e:
            print(f"Не удалось открыть файл: {e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
