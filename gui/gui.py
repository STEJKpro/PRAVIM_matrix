from matrix_generator import excel_generator
import customtkinter
from tkinter.filedialog import asksaveasfile
import pyperclip

class TopWindow:
    def __init__(self, root, lable_text):
        self.root = root
        self.window = customtkinter.CTkToplevel()
        self.window.title('Сообщение')
        # self.window.geometry ("300x100")
        self.window.geometry("%dx%d+%d+%d" % (300, 100, root.winfo_x() + 700/4, root.winfo_y() + 550/4))
         
        label = customtkinter.CTkLabel(self.window, text=lable_text)
        label.pack()
         
        btn = customtkinter.CTkButton(self.window, text='Окей', command=self.close)
        btn.pack(anchor='s')
        self.root.withdraw()
    def close(self):
        self.root.deiconify()
        self.window.destroy()
        

class App(customtkinter.CTk):

    def keypress(self, e):
        if e.keycode == 86 and e.keysym != 'v':
            e.widget.event_generate("<<Paste>>")
            
        elif e.keycode == 67 and e.keysym != 'c':
            e.widget.event_generate("<<Copy>>")
            
        elif e.keycode == 88 and e.keysym != 'x':
            e.widget.event_generate("<<Сut>>")
            
        elif e.keycode == 65 and e.keysym != 'x':
            e.widget.event_generate("<<SelectAll>>")
            
    def get_signature_text_from_file(self,) -> str:
        with open('gui\src\signature_text.txt', 'r', encoding='utf-8') as file:
            return file.read()
             
    def update_signature_text_to_file(self, signature) -> None:
        with open('gui\src\signature_text.txt', 'w', encoding='utf-8') as file:
            file.write(signature)
      
    def __init__(self):
        super().__init__()
        self.title("Генератор планограммы (ассортиментной матрицы)")
        self.geometry("700x750")

        # set grid layout 1x2
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        customtkinter.set_appearance_mode('dark')
        self.bind("<Control-KeyPress>", self.keypress)
        # self.bind("<Control-86>", lambda: print('Okkkkkk'))

        


        # create home frame
        self.home_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.home_frame.grid_columnconfigure(0, weight=1)



        self.home_frame_large_image_label = customtkinter.CTkLabel(self.home_frame, text="Генератор планограмм для Haiba/Cron",)
        self.home_frame_large_image_label.grid(row=0, column=0, padx=20, pady=10)
        
        #Add textbox
        self.textbox = customtkinter.CTkTextbox(master=self.home_frame, width=400, height=150, corner_radius=0, font = ('',16),)
        self.textbox.grid(row=0, column=0, sticky="nsew", padx=20, pady=10)
        self.textbox.insert("0.0", "HB101, HB101, HB101, HB10590-7, HB10803-7, HB10505-3\nHB101, HB101, HB101, HB60590-7, HB22803-7, HB60505-3\nHB10505, HB10505-8, HB10198, HB10590, CN10523, CN10129\nHB60505, HB60505-8, HB60198, HB60590, CN60523, CN22129")
        
        #button to clear textbox
        self.clear_button = customtkinter.CTkButton(master=self.home_frame, text="Очистить поле", command=lambda: self.textbox.delete('0.0','end'))
        self.clear_button.grid(row=0, column=0, padx=20, pady=10, sticky='ne' )
        
        
        self.slider = customtkinter.CTkSlider(master=self.home_frame, from_=125, to=195, command=self.set_slider_lable_text, number_of_steps=14)
        self.slider.grid(row=2, column=0, sticky="nwe", padx=10, pady=10)
        
        self.slider_label = customtkinter.CTkLabel(
            self.home_frame,
            text=f"Размер изображений\n135 - смесители, 160 - ДП.\n Текущее значение: {int(self.slider.get())}",
            justify='left',
            font=('', 16),
        )
        self.slider_label.grid(row=1, column=0, sticky="sw", padx=10, pady=10, ) 


        self.customer_field = customtkinter.CTkEntry(self.home_frame, placeholder_text="Контрагент", font=('', 20))
        self.customer_field.grid(row=3, column=0, sticky="nsew", padx=10, pady=20, ) 
        
        self.signature_switch = customtkinter.CTkSwitch(
            self.home_frame,
            text="Добавить подпись\n ** - Выделить жирным",
            font=('', 16),
            command=lambda: print(self.signature_switch.get())
        )
        self.signature_switch.grid(row=4, column=0, sticky="nsew", padx=10, pady=(20,)) 
        
        
        self.signature_field = customtkinter.CTkTextbox(self.home_frame, height=120, font =('', 16))
        self.signature_field.grid(row=5, column=0, sticky="nsew", padx=10, pady=(0,20) )
        self.signature_field.insert('0.0', str(self.get_signature_text_from_file()))
        
        self.save_signature_button = customtkinter.CTkButton(
            master=self.home_frame, 
            text="Сохранить подпись", 
            command=lambda: self.update_signature_text_to_file(self.signature_field.get("1.0", 'end-1c').strip())
        )
        self.save_signature_button.grid(row=5, column=0, padx=20, pady=0, sticky='ne' )
        
        
        self.update_switch = customtkinter.CTkSwitch(
            self.home_frame,
            text="Обновить прайс. \n(Если включено, то скачает прайс-лист из облака. \nИначе - использует ранее скачанные)",
            font=('', 16),
        )
        self.update_switch.grid(row=6, column=0, sticky="nsew", padx=10, pady=20, ) 
        
        self.home_frame_button_1 = customtkinter.CTkButton(self.home_frame, text="Сгенерировать и сохранить!", height=50, font=('', 26), command= lambda: self.save_file())
        self.home_frame_button_1.grid(row=7, column=0, padx=20, pady=10, ipadx = 20, ipady = 10)
        
        self.home_frame.grid(row=0, column=1, sticky="nsew")


        
    def set_slider_lable_text(self, value):
        self.slider_label.configure(text = f"Размер изображений\n135 - смесители, 160 - ДП\n Текущее значение: {int(value)}")
    
    def parse_matrix (self, string) -> list:
        matrix = []
        for line in string.split('\n'):
                matrix.append ([x.strip() for x in line.strip().split(',')])
        return matrix

    
    def save_file(self):
        try:
            f = asksaveasfile(initialfile = 'Файл_без_имени.xlsx',
                defaultextension=".xlsx",filetypes=[("All Files","*.*"),("Excel files","*.xlsx")], mode='wb')
            f.write(
                excel_generator.generate_excel_file(
                    matrix= self.parse_matrix(self.textbox.get("1.0", 'end-1c').strip()),
                    img_size=int(self.slider.get()),
                    customer=self.customer_field.get(),
                    signature_text = self.signature_field.get("1.0", 'end-1c').strip() if self.signature_switch.get() == 1 else None,
                    get_updates = int(self.update_switch.get()),
                )
            )
            TopWindow(self, 'Файл успешно сохранен!')
        except Exception as ex:
            TopWindow(self, f'Произошла ошибка:\n {ex}')
        finally:
            f.close
            

if __name__ == "__main__":
    app = App()
    app.mainloop()
