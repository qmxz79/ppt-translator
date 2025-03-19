from tkinter import Tk, Button, filedialog, messagebox, ttk, Toplevel, StringVar, Radiobutton
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
            messagebox.showinfo("文件保存", f"文件将保存至: {os.path.basename(self.output_file)}")

    def translate(self):
        if self.input_file and self.output_file:
            self.progress["value"] = 0
            self.master.update_idletasks()
            try:
                set_translation_engine(self.translation_engine.get())
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

        Radiobutton(settings_window, text="DeepSeek", variable=self.translation_engine, value="DeepSeek").pack(anchor="w")
        Radiobutton(settings_window, text="国内大模型API", variable=self.translation_engine, value="DomesticAPI").pack(anchor="w")
        Radiobutton(settings_window, text="本地ollama部署", variable=self.translation_engine, value="LocalOllama").pack(anchor="w")

if __name__ == "__main__":
    root = Tk()
    gui = PPTTranslatorGUI(root)
    root.mainloop()