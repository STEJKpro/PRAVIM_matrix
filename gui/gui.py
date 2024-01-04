# import customtkinter
# import os
    

# class App(customtkinter.CTk):
#     def __init__(self):
#         super().__init__()

#         # configure window
#         self.title("Генератор планограм")
#         self.geometry(f"{1100}x{580}")
        
#         # configure grid layout (4x4)
#         self.grid_columnconfigure(1, weight=1)
#         self.grid_columnconfigure((2, 3), weight=0)
#         self.grid_rowconfigure((0, 1, 2), weight=1)
        
#         self.label = customtkinter.CTkLabel(self, text="Генератор планограм", fg_color="transparent")
#         self.label.grid(row=0, column=0, rowspan=1, sticky="nsew")
#         self.label.grid_rowconfigure(4, weight=1)


#         # create home frame
#         self.home_frame_button_1 = customtkinter.CTkButton(self, text=" Сгенерировать",)
#         self.home_frame_button_1.grid(row=1, column=0, padx=20, pady=10)


# if __name__ == "__main__":
#     customtkinter.set_appearance_mode("dark")
#     app = App()
    # app.mainloop()
from matrix_generator import excel_generator
import customtkinter
from tkinter.filedialog import asksaveasfile


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("image_example.py")
        self.geometry("700x450")

        # set grid layout 1x2
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        customtkinter.set_appearance_mode('dark')


        # create home frame
        self.home_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.home_frame.grid_columnconfigure(0, weight=1)

        self.home_frame_large_image_label = customtkinter.CTkLabel(self.home_frame, text="Генератор планограмм для Haiba/Cron",)
        self.home_frame_large_image_label.grid(row=0, column=0, padx=20, pady=10)
        
        #Add textbox
        self.textbox = customtkinter.CTkTextbox(master=self.home_frame, width=400, corner_radius=0, font = ('',20))
        self.textbox.grid(row=0, column=0, sticky="nsew", padx=20, pady=10)
        self.textbox.insert("0.0", "HB101, HB101, HB101, HB10590-7, HB10803-7, HB10505-3\nHB101, HB101, HB101, HB60590-7, HB22803-7, HB60505-3\nHB10505, HB10505-8, HB10198, HB10590, CN10523, CN10129\nHB60505, HB60505-8, HB60198, HB60590, CN60523, CN22129")
        
        self.slider = customtkinter.CTkSlider(master=self.home_frame, from_=100, to=300, command=self.set_slider_lable_text, number_of_steps=40)
        self.slider.grid(row=2, column=0, sticky="nwe", padx=10, pady=10)
        
        self.slider_label = customtkinter.CTkLabel(self.home_frame, text=f"Размер изображений\n135 - смесители, 160 - ДП.\n Текущее значение: {int(self.slider.get())}", justify='left')
        self.slider_label.grid(row=1, column=0, sticky="sw", padx=10, pady=10, ) 

        # slider = customtkinter.CTkSlider(app, from_=0, to=100, command=slider_event)
        
        
        self.home_frame_button_1 = customtkinter.CTkButton(self.home_frame, text="Сгенрировать!", height=50, font=('', 32), command= lambda: save_file())
        self.home_frame_button_1.grid(row=3, column=0, padx=20, pady=10, ipadx = 20, ipady = 10)
        self.home_frame.grid(row=0, column=1, sticky="nsew")

    def set_slider_lable_text(self, value):
        self.slider_label.configure(text = f"Размер изображений\n135 - смесители, 160 - ДП\n Текущее значение: {int(value)}")
    
def save_file():
    f = asksaveasfile(initialfile = 'Файл_без_имени.xlsx',
        defaultextension=".xlsx",filetypes=[("All Files","*.*"),("Excel files","*.xlsx")], mode='wb')
    f.write(excel_generator.generate_excel_file())
    f.close
        
if __name__ == "__main__":
    app = App()
    app.mainloop()
