import tkinter as tk

root = tk.Tk()
root.title("Exemplo de Cores Azuis")

# Aplicando as cores em rótulos
label1 = tk.Label(root, text="Azul mais claro (#1E4CA3)", fg="#1E4CA3", font=("Arial", 14))
label1.pack(pady=10)

label2 = tk.Label(root, text="Azul médio (#2A5DB0)", fg="#2A5DB0", font=("Arial", 14))
label2.pack(pady=10)

label3 = tk.Label(root, text="Azul claro (#3B6FCF)", fg="#3B6FCF", font=("Arial", 14))
label3.pack(pady=10)

root.mainloop()