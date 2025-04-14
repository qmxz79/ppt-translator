from tkinter import Tk, Button, filedialog, messagebox, ttk, Toplevel, StringVar, Radiobutton, OptionMenu
import os
from pptx import Presentation
from cn2enppt import translate_ppt, set_translation_engine  # Assuming translate_ppt and set_translation_engine are in cn2enppt.py

class PPTTranslatorGUI:
    def __init__(self, master):
        self.master = master
        master.title("PPT Translator")

        self.label = Button(master, text="选择要翻译的课件", command=self.load_file)
        self.label.pack(pady=10)

        self.save_button = Button(master, text="另存为", command=self.save_file)
        self.save_button.pack(pady=10)

        self.translate_button = Button(master, text="转换", command=self.translate)
        self.translate_button.pack(pady=10)

        self.settings_button = Button(master, text="设置", command=self.open_settings)
        self.settings_button.pack(pady=10)

        self.progress = ttk.Progressbar(master, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=10)

        self.input_file = None
        self.output_file = None
        self.translation_engine = StringVar(value="DeepSeek")
        self.model_option = StringVar(value="")

    def load_file(self):
        self.input_file = filedialog.askopenfilename(
            title="选择PPT文件",
            filetypes=[("PowerPoint files", "*.pptx")]
        )
        if self.input_file:
            messagebox.showinfo("文件选择", f"选择的文件: {os.path.basename(self.input_file)}")

    def save_file(self):
        self.output_file = filedialog.asksaveasfilename(
            title="另存为",
            defaultextension=".pptx",
            filetypes=[("PowerPoint files", "*.pptx")]
        )
        if self.output_file:
            if os.path.exists(self.output_file):
                if not messagebox.askyesno("文件已存在", f"{os.path.basename(self.output_file)} 已存在。是否覆盖？"):
                    return
            messagebox.showinfo("文件保存", f"文件将保存至: {os.path.basename(self.output_file)}")

    def translate(self):
        if self.input_file and self.output_file:
            self.progress["value"] = 0
            self.master.update_idletasks()
            try:
                set_translation_engine(self.translation_engine.get(), self.model_option.get())
                translate_ppt(self.input_file, self.output_file, self.update_progress)
                messagebox.showinfo("翻译完成", f"翻译结果已保存至: {os.path.basename(self.output_file)}")
            except Exception as e:
                messagebox.showerror("错误", f"翻译失败: {e}")
        else:
            messagebox.showwarning("警告", "请先选择要翻译的课件文件和保存位置。")

    def update_progress(self, value):
        self.progress["value"] = value
        self.master.update_idletasks()

    def open_settings(self):
        settings_window = Toplevel(self.master)
        settings_window.title("设置")

        Radiobutton(settings_window, text="DeepSeek", variable=self.translation_engine, value="DeepSeek", command=self.update_model_options).pack(anchor="w")
        Radiobutton(settings_window, text="国内大模型API", variable=self.translation_engine, value="DomesticAPI", command=self.update_model_options).pack(anchor="w")
        Radiobutton(settings_window, text="本地ollama部署", variable=self.translation_engine, value="LocalOllama", command=self.update_model_options).pack(anchor="w")

        self.model_option_menu = OptionMenu(settings_window, self.model_option, "")
        self.model_option_menu.pack(anchor="w")

        Button(settings_window, text="确定", command=settings_window.destroy).pack(pady=10)

    def update_model_options(self):
        engine = self.translation_engine.get()
        if engine == "DeepSeek":
            options = ["deepseek-v1", "deepseek-r1"]
        elif engine == "DomesticAPI":
            options = ["Domestic Model 1", "Domestic Model 2"]
        elif engine == "LocalOllama":
            options = ["deepseek-r1:1.5b", "huihui_ai/deepseek-r1-abliterated:7b"]
        else:
            options = []

        self.model_option.set(options[0] if options else "")
        menu = self.model_option_menu["menu"]
        menu.delete(0, "end")
        for option in options:
            menu.add_command(label=option, command=lambda value=option: self.model_option.set(value))

if __name__ == "__main__":
    root = Tk()
    gui = PPTTranslatorGUI(root)
    root.mainloop()