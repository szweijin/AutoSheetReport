import tkinter as tk
from tkinter import messagebox
from report_generator import generate_report


def run_report():
    status_var.set("執行中...")
    root.update()
    success, msg = generate_report()
    status_var.set(msg)
    if success:
        messagebox.showinfo("成功", msg)
    else:
        messagebox.showerror("錯誤", msg)

root = tk.Tk()
root.title("Google Sheet 報表工具")
root.geometry("400x200")

btn = tk.Button(root, text="產出部門報表", font=("Arial", 14), command=run_report)
btn.pack(pady=20)

status_var = tk.StringVar()
status_label = tk.Label(root, textvariable=status_var, fg="blue", wraplength=380)
status_label.pack(pady=10)

exit_btn = tk.Button(root, text="結束程式", command=root.quit)
exit_btn.pack()

root.mainloop()
