import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import subprocess

def select_file():
    file_path = filedialog.askopenfilename(
        title="Оберіть файл",
        filetypes=(("Excel файли", "*.xlsx"), ("Усі файли", "*.*"))
    )
    if file_path:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)

def generate_tree():
    excel_file = entry_file_path.get()
    if not excel_file:
        messagebox.showerror("Помилка", "Будь ласка, оберіть Excel файл!")
        return

    if not os.path.exists(excel_file):
        messagebox.showerror("Помилка", "Файл не існує!")
        return

    try:
        subprocess.run(["python3", "total_family_tree_plotter.py"], check=True)
        messagebox.showinfo("Успіх", "Дерево згенеровано успішно!")
    except Exception as e:
        messagebox.showerror("Помилка", f"Сталася помилка під час генерації: {str(e)}")

# Головне вікно
root = tk.Tk()
root.title("Генератор родинного дерева")
root.state("zoomed")  # Відкривається у повноекранному режимі
root.configure(bg="#f0f4f7")  # Фон із м'яким кольором

# Задаємо стиль
style = ttk.Style(root)
style.theme_use("clam")
style.configure("TButton", font=("Arial", 16, "bold"), padding=10)
style.configure("TLabel", font=("Arial", 14))
style.configure("TEntry", font=("Arial", 12), padding=5)

# Заголовок
label_title = ttk.Label(root, text="ГЕНЕРАЦІЯ РОДИННИХ ДЕРЕВ", font=("Arial", 40, "bold"), background="#f0f4f7")
label_title.pack(pady=50)

# Поле для вибору файлу (в центрі)
frame_file = ttk.Frame(root, padding="20 20 20 20")
frame_file.pack(pady=50)

entry_file_path = ttk.Entry(frame_file, width=30, font=("Arial", 14))
entry_file_path.pack(side=tk.LEFT, padx=(0, 10))

btn_browse = ttk.Button(frame_file, text="Обрати файл", command=select_file)
btn_browse.pack(side=tk.LEFT)

# Кнопка "Згенерувати" (велика і в центрі)
btn_generate = ttk.Button(root, text="ЗГЕНЕРУВАТИ", command=generate_tree)
btn_generate.pack(pady=50, ipadx=50, ipady=20)

# Інструкція внизу
label_footer = ttk.Label(root, text="* Оберіть файл з інформацією про родичів, потім натисніть 'Згенерувати'", 
                         foreground="gray", font=("Arial", 12), background="#f0f4f7")
label_footer.pack(side=tk.BOTTOM, pady=20)

# Запуск програми
root.mainloop()
