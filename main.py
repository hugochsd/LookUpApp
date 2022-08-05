import tkinter
import tkinter.messagebox
import customtkinter
from tkinter import filedialog
import pandas as pd
import time
import requests

customtkinter.set_appearance_mode("Light")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):

    WIDTH = 780
    HEIGHT = 520

    def __init__(self):
        super().__init__()

        self.title("SecOps Tool")
        self.geometry(f"{App.WIDTH}x{App.HEIGHT}")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)  # call .on_closing() when app gets closed

        # ============ create two frames ============

        # configure grid layout (2x1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.frame_left = customtkinter.CTkFrame(master=self,
                                                 width=180,
                                                 corner_radius=0)
        self.frame_left.grid(row=0, column=0, sticky="nswe")

        self.frame_right = customtkinter.CTkFrame(master=self)
        self.frame_right.grid(row=0, column=1, sticky="nswe", padx=20, pady=20)

        # ============ frame_left ============

        # configure grid layout (1x11)
        self.frame_left.grid_rowconfigure(0, minsize=10)   # empty row with minsize as spacing
        self.frame_left.grid_rowconfigure(5, weight=1)  # empty row as spacing
        self.frame_left.grid_rowconfigure(8, minsize=20)    # empty row with minsize as spacing
        self.frame_left.grid_rowconfigure(11, minsize=10)  # empty row with minsize as spacing

        self.label_1 = customtkinter.CTkLabel(master=self.frame_left,
                                              text="SecOps LookUp Tool",
                                              text_font=("Roboto Medium", 18))  # font name and size in px
        self.label_1.grid(row=1, column=0, pady=10, padx=10)

        self.button_1 = customtkinter.CTkButton(master=self.frame_left,
                                                text="IPv4 LookUp",
                                                command=self.buttonIP_event)
        self.button_1.grid(row=2, column=0, pady=10, padx=20)

        self.button_2 = customtkinter.CTkButton(master=self.frame_left,
                                                text="URL LookUp",
                                                command=self.buttonURL_event)
        self.button_2.grid(row=3, column=0, pady=10, padx=20)
        self.label_info_2 = customtkinter.CTkLabel(master=self.frame_left,
                                                   text="Requirements :    \n" +
                                                        "\n" +
                                                        "IPv4 - Column label ->           'SenderIPv4'   \n" +
                                                        "          - Sheet name  label -> 'infected_ips' \n" +
                                                        "\n" +
                                                        "URL  - Column  label ->          'URL' \n" +
                                                        "          - Sheet name  label -> 'URL' \n",
                                                   text_font=("Roboto Medium",10),
                                                   height=100,
                                                   corner_radius=6,  # <- custom corner radius
                                                   fg_color=("white", "gray38"),  # <- custom tuple-color
                                                   justify=tkinter.LEFT)
        self.label_info_2.grid(column=0, row=5, sticky="nwe", padx=15, pady=15)

        self.label_mode = customtkinter.CTkLabel(master=self.frame_left, text="Appearance Mode:")
        self.label_mode.grid(row=9, column=0, pady=0, padx=20, sticky="w")

        self.optionmenu_1 = customtkinter.CTkOptionMenu(master=self.frame_left,
                                                        values=["Light", "Dark", "System"],
                                                        command=self.change_appearance_mode)
        self.optionmenu_1.grid(row=10, column=0, pady=10, padx=20, sticky="w")

        # ============ frame_right ============

        # configure grid layout (3x7)
        self.frame_right.rowconfigure((0, 1, 2, 3), weight=1)
        self.frame_right.rowconfigure(7, weight=10)
        self.frame_right.columnconfigure((0, 1), weight=1)
        self.frame_right.columnconfigure(2, weight=0)

        self.frame_info = customtkinter.CTkFrame(master=self.frame_right)
        self.frame_info.grid(row=0, column=0, columnspan=2, rowspan=4, pady=20, padx=20, sticky="nsew")

        # ============ frame_info ============

        # configure grid layout (1x1)
        self.frame_info.rowconfigure(0, weight=1)
        self.frame_info.columnconfigure(0, weight=1)

        self.label_info_1 = customtkinter.CTkLabel(master=self.frame_info,
                                                   text="Using APIVoid For IP and URL   \n" +
                                                        "https://docs.apivoid.com/      \n" +
                                                        "See requirements                 " ,
                                                   height=100,
                                                   corner_radius=6,  # <- custom corner radius
                                                   fg_color=("white", "gray38"),  # <- custom tuple-color
                                                   justify=tkinter.LEFT)
        self.label_info_1.grid(column=0, row=0, sticky="nwe", padx=15, pady=15)
        self.progressbar = customtkinter.CTkProgressBar(master=self.frame_info)
        self.progressbar.grid(row=1, column=0, sticky="ew", padx=15, pady=15)

        # ============ frame_right ============

        self.radio_var = tkinter.IntVar(value=0)

        self.label_radio_group = customtkinter.CTkLabel(master=self.frame_right,
                                                        text="Filter:",
                                                        text_font=("Roboto Medium", 16))
        self.label_radio_group.grid(row=0, column=2, columnspan=1, pady=20, padx=10, sticky="")

        self.radio_button_1 = customtkinter.CTkRadioButton(master=self.frame_right,
                                                           variable=self.radio_var,
                                                           text="Cleared (.xslx)",
                                                           text_font=("Roboto Medium", -14),
                                                           value=0)
        self.radio_button_1.grid(row=1, column=2, pady=10, padx=20, sticky="n")

        self.radio_button_2 = customtkinter.CTkRadioButton(master=self.frame_right,
                                                           variable=self.radio_var,
                                                           text="     Risky (.xlsx)",
                                                           text_font=("Roboto Medium", -14),
                                                           value=1)
        self.radio_button_2.grid(row=2, column=2, pady=10, padx=20, sticky="n")

        self.radio_button_3 = customtkinter.CTkRadioButton(master=self.frame_right,
                                                           variable=self.radio_var,
                                                           text="     Report (.txt)",
                                                           text_font=("Roboto Medium", -14),
                                                           value=2)
        self.radio_button_3.grid(row=3, column=2, pady=10, padx=20, sticky="n")




        self.slider_2 = customtkinter.CTkSlider(master=self.frame_right,
                                                command=self.progressbar.set)
        self.slider_2.grid(row=5, column=0, columnspan=2, pady=10, padx=20, sticky="we")


        self.entry = customtkinter.CTkEntry(master=self.frame_right,
                                            width=120,
                                            placeholder_text="Search")
        self.entry.grid(row=8, column=0, columnspan=2, pady=20, padx=20, sticky="we")

        self.button_5 = customtkinter.CTkButton(master=self.frame_right,
                                                text="Submit",
                                                border_width=2,  # <- custom border_width
                                                fg_color=None,  # <- no fg_color
                                                command=self.button_event)
        self.button_5.grid(row=8, column=2, columnspan=1, pady=20, padx=20, sticky="we")

        # set default values
        self.optionmenu_1.set("Dark")
        self.radio_button_1.select()
        self.progressbar.set(0)

    def button_event(self):
        print("Button pressed")

    def buttonIP_event(self):
        apivoid_key = '902345018587a33f9c99134d8011a6739a1a7873'
        self.filename = filedialog.askopenfilename(initialdir="C:/", title="Select A File", filetypes=(("Excel (.xlsx)", "*.xlsx"), ("All Files", "*.*")))
        IPdf = pd.read_excel(self.filename, "infected_ips")
        IP_list = IPdf['SenderIPv4'].values.tolist()

        goodIPs = []
        badIPs = []

        if self.radio_button_3.check_state == True:
            # open .txt to write report
            outfile = open('Report.txt', 'w')
            for ip in IP_list:
                url = f'https://endpoint.apivoid.com/iprep/v1/pay-as-you-go/?key={apivoid_key}&ip={ip}'
                r = requests.get(url)
                outfile.write('\n#####################################################################\n')
                outfile.write(f'IP : {r.json()["data"]["report"]["ip"]}\n')
                outfile.write(f'DETECTIONS : {r.json()["data"]["report"]["blacklists"]["detections"]}\n')
                outfile.write(f'RISK SCORE : {r.json()["data"]["report"]["risk_score"]["result"]}\n')
                outfile.write(f'COUNTRY : {r.json()["data"]["report"]["information"]["country_name"]}\n')
                outfile.write(f'CITY : {r.json()["data"]["report"]["information"]["city_name"]}\n')
                outfile.write(f'REGION : {r.json()["data"]["report"]["information"]["region_name"]}\n')
                outfile.write(f'ISP : {r.json()["data"]["report"]["information"]["isp"]}\n')
                outfile.write(f'REVERSE DNS : {r.json()["data"]["report"]["information"]["reverse_dns"]}\n')
                outfile.write(f'ANONIMITY : {r.json()["data"]["report"]["anonymity"]}\n')
                print(IP_list.index(ip))
            self.progressbar.set(1)
            outfile.close()

        elif self.radio_button_1.check_state == True:
            for ip in IP_list:
                #time.sleep(0.5)
                url = f'https://endpoint.apivoid.com/iprep/v1/pay-as-you-go/?key={apivoid_key}&ip={ip}'
                r = requests.get(url)
                if r.json()['data']['report']['blacklists']['detections'] == 0:
                    goodIPs.append(ip)

            goodIPdata = pd.DataFrame({'Cleared IPs': goodIPs})
            writer = pd.ExcelWriter('Cleared.xlsx')
            goodIPdata.to_excel(writer, sheet_name='clearedIPsheet')
            worksheet = writer.sheets['clearedIPsheet']
            worksheet.set_column(1, 1, 50)
            writer.save()
            self.progressbar.set(1)
        elif self.radio_button_2.check_state == True:
            for ip in IP_list:
                #time.sleep(0.5)
                url = f'https://endpoint.apivoid.com/iprep/v1/pay-as-you-go/?key={apivoid_key}&ip={ip}'
                r = requests.get(url)
                if r.json()['data']['report']['blacklists']['detections'] > 0:
                    badIPs.append(ip)

            badIPdata = pd.DataFrame({'Risky IPs': badIPs})
            writer = pd.ExcelWriter('Risky.xlsx')
            badIPdata.to_excel(writer, sheet_name='RiskyIPsheet')
            worksheet = writer.sheets['RiskyIPsheet']
            worksheet.set_column(1, 1, 50)
            writer.save()
            self.progressbar.set(1)

    def buttonURL_event(self):
        apivoid_key = '902345018587a33f9c99134d8011a6739a1a7873'
        self.filename = filedialog.askopenfilename(initialdir="C:/", title="Select A File", filetypes=(("Excel (.xlsx)", "*.xlsx"), ("All Files", "*.*")))
        URLdf = pd.read_excel(self.filename, "URL")
        URL_list = URLdf['URL'].values.tolist()

        goodURLs = []
        badURLs = []

        if self.radio_button_3.check_state:
            # open .txt to write report
            outfile = open('Report.txt', 'w')
            for url in URL_list:
                #time.sleep(0.5)
                api = f'https://endpoint.apivoid.com/urlrep/v1/pay-as-you-go/?key={apivoid_key}&url={url}'
                r = requests.get(api)
                outfile.write('\n#####################################################################\n')
                outfile.write(f'COUNTRY CODE : {r.json()["data"]["report"]["server_details"]["country_code"]}\n')
                outfile.write(f'ISP : {r.json()["data"]["report"]["server_details"]["isp"]}\n')
                outfile.write(f'IP: {r.json()["data"]["report"]["server_details"]["ip"]}\n')
                outfile.write(f'DETECTIONS: {r.json()["data"]["report"]["domain_blacklist"]["detections"]}\n')
                outfile.write(f'FILE TYPE: {r.json()["data"]["report"]["file_type"]}\n')
                print(URL_list.index(url))
            self.progressbar.set(1)
            outfile.close()
        elif self.radio_button_1.check_state:
            for URL in URL_list:
                #time.sleep(0.5)
                url = f'https://endpoint.apivoid.com/urlrep/v1/pay-as-you-go/?key={apivoid_key}&url={URL}'
                r = requests.get(url)
                if r.json()['data']['report']['domain_blacklist']['detections'] == 0:
                    goodURLs.append(URL)

            goodURLdata = pd.DataFrame({'Cleared URLs': goodURLs})
            writer = pd.ExcelWriter('Cleared.xlsx')
            goodURLdata.to_excel(writer, sheet_name='clearedURLsheet')
            worksheet = writer.sheets['clearedURLsheet']
            worksheet.set_column(1, 1, 50)
            writer.save()
            self.progressbar.set(1)
        elif self.radio_button_2.check_state:
            for url in URL_list:
                #time.sleep(0.5)
                api = f'https://endpoint.apivoid.com/urlrep/v1/pay-as-you-go/?key={apivoid_key}&url={url}'
                r = requests.get(api)
                if r.json()['data']['report']['domain_blacklist']['detections'] > 0:
                    badURLs.append(url)

            badURLdata = pd.DataFrame({'Risky URLs': badURLs})
            writer = pd.ExcelWriter('Risky.xlsx')
            badURLdata.to_excel(writer, sheet_name='RiskyURLsheet')
            worksheet = writer.sheets['RiskyURLsheet']
            worksheet.set_column(1, 1, 50)
            writer.save()
            self.progressbar.set(1)

    def change_appearance_mode(self, new_appearance_mode):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def on_closing(self, event=0):
        self.destroy()

if __name__ == "__main__":
    app = App()
    app.mainloop()