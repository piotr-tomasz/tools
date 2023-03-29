
import tkinter as tk
import tkinter.ttk as ttk


def display_tk_window(measurements,window_size='500x800'):
    def dict_to_treeview(data, treeview=None, parent='', child=''):
        if treeview is None:
            treeview = ttk.Treeview()
        for k, v in data.items():
            item = treeview.insert(parent, 'end', text=k)
            if isinstance(v, dict):
                dict_to_treeview(v, treeview, item, '')
            else:
                treeview.insert(item, 'end', text=v)
        return treeview

    def close_window():
        root.destroy()

    root = tk.Tk()
    treeview = dict_to_treeview(measurements)
    treeview.pack(fill='both', expand=True)
    width = treeview.winfo_reqwidth()
    height = treeview.winfo_reqheight()
    root.geometry(f'{width}x{height}')
    root.geometry(window_size)
    exit_button = tk.Button(root, text="Exit", command=close_window)
    exit_button.pack()
    root.mainloop()