from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from PIL import Image as PImage, ImageTk as PImageTk
import matplotlib.font_manager as mplfm
from tkinter import messagebox as msg
import matplotlib.pyplot as plt
import tkintermapview as tkmap
from tkinter import filedialog
from customtkinter import *
from tkinter import *
import pandas as pd
import configparser
import os

# make sure this directory contain errchk.py
try:
    from errchk import *
except ImportError as e:
    msg.showerror('Error', f'{e}')

# make sure program can be run in any directory
program_path = os.path.dirname(os.path.abspath(__file__))

# Config file and path
config = configparser.ConfigParser(allow_no_value=True)
err = check_error()['err']
err_code = list(err.keys())[0]
err_info = list(err.values())[0]
if err_code == '404':
    msg.showwarning(
        'พบข้อผิดพลาด', f'ไฟล์โปรแกรมบางส่วนสูญหาย\nกดตกลงเพื่อทำการซ่อมแซม\nไฟล์: {err_info}')
elif err_code == '500':
    msg.showerror(
        'พบข้อผิดพลาด',
        f'เกิดความเสียหายอย่างหนักกับไฟล์ที่ใช้ในโปรแกรม\nโปรดติดตั้งใหม่หรือติดต่อนักพัฒนา\nไฟล์: {err_info}')
else:
    config.read(os.path.join(program_path, 'config\\cfg.ini'))
# Config file and path

# font for matplotlib
mplfm.fontManager.addfont(os.path.join(program_path, 'config\\font\\NotoSerifThai.ttf'))
plt.rc('font', family='Noto Serif Thai')
# font for matplotlib

# Theme and mode for application GUI
set_default_color_theme('dark-blue')
set_appearance_mode(config.get('THEME', 'mode'))
# set list for combobox
theme = []
if config.get('THEME', 'mode') == 'light':
    theme = ['สว่าง', 'มืด']
else:
    theme = ['มืด', 'สว่าง']

"""fuction section"""


# close window
def on_closing():
    os._exit(0)


# font for tkinter
def myfont(size=16, weight='normal'):
    return 'Noto Serif Thai', size, weight


# change theme
# color dictionary for graph
color_dict = {'colorset': {'0': ['#ff4d00', '#ff9900', '#ffe600', '#e6ff00', '#99ff00', '#4dff00', '#00ff00',
                                 '#00ff4d', '#00ff99', '#00ffe6', '#00e6ff', '#0099ff', '#004dff', '#0000ff'],
                           '1': ['#ff0000', '#ff4d00', '#ff9900', '#ffe600', '#e6ff00', '#99ff00', '#4dff00',
                                 '#00ff00', '#00ff4d', '#00ff99', '#00ffe6', '#00e6ff', '#0099ff', '#004dff'],
                           '2': ['#00FFF6', '#00E6FF', '#00CCFF', '#00B2FF', '#0099FF', '#0080FF', '#0066FF',
                                 '#004DFF', '#0033FF', '#0019FF', '#0000FF', '#1900FF', '#3300FF', '#4D00FF'],
                           '3': ['#FD841F', '#FFA500', '#FFC000', '#FFD700', '#FFE600', '#FFFF00', '#E6FF00',
                                 '#CCFF00', '#B2FF00', '#99FF00', '#80FF00', '#66FF00', '#4DFF00', '#33FF00'],
                           '4': plt.rcParams['axes.prop_cycle'].by_key()['color']  # default color of matplotlib
                           },
              'graphbg': {'0': '#ffffff', '1': '#e0e0e0'},  # background color of graph
              }


# change theme function
def change_theme(new_theme):
    theme_dict = {'สว่าง': 'light', 'มืด': 'dark'}
    set_appearance_mode(theme_dict[new_theme])
    config['THEME'] = {'mode': theme_dict[new_theme]}  # read from config
    # write to cfg.ini
    with open(os.path.join(program_path, 'config\\cfg.ini'), 'w') as f:
        config.write(f)


"""fuction section"""

# main window
root = CTk()
root.title('ข้อมูลผลผลิตข้าวนาปรังแยกตามจังหวัด')
root.iconbitmap(os.path.join(program_path, 'img\\icon.ico'))
root.protocol("WM_DELETE_WINDOW", on_closing)


def center_screen(app, w, h):
    # setup size of the window and move it to the center of the screen
    screen_width = app.winfo_screenwidth()  # width of the screen
    screen_height = app.winfo_screenheight()  # height of the screen
    x = int((screen_width / 2) - (w / 2))  # calculate x for the Tk root window
    y = int((screen_height / 2) - (h / 2))  # calculate y for the Tk root window
    # setup size of the window and move it to the center of the screen
    return "{}x{}+{}+{}".format(w, h, x, y - 40)


root.geometry(center_screen(root, 1600, 900))

# config grid of row and column for root
root.grid_columnconfigure(1, weight=1)
root.grid_rowconfigure(0, weight=1)


# load image
def load_image(path, image_size):
    return PImageTk.PhotoImage(PImage.open(os.path.join(program_path, path)).resize((image_size, image_size)))


img_chart_histogram = load_image("img\\chart-histogram.png", 20)
img_settings_image = load_image("img\\settings.png", 20)
img_chart_pie = load_image("img\\chart-pie.png", 20)
img_document = load_image("img\\document.png", 20)
img_download = load_image("img\\download.png", 20)
img_refresh = load_image("img\\refresh.png", 20)
img_comp = load_image("img\\compare.png", 20)
img_stats = load_image("img\\stats.png", 20)
img_user = load_image("img\\user.png", 20)

"""frame section"""
# create frame
frame_left = CTkFrame(root, width=180, corner_radius=0)
frame_left.grid(row=0, column=0, sticky="nsew")
frame_right = CTkFrame(root)
frame_right.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
frame_info = CTkFrame(frame_right)
frame_info.grid(row=0, column=0, columnspan=2, pady=5, padx=5, sticky="nsew")

# configure frame
frame_right.rowconfigure(7, weight=10)
frame_right.columnconfigure(1, weight=1)
frame_right.columnconfigure(2, weight=0)
frame_info.rowconfigure(0, weight=1)
frame_info.columnconfigure(0, weight=1)

# empty row with minsize as spacing
frame_left.grid_rowconfigure(0, minsize=10)
frame_left.grid_rowconfigure(5, weight=1)
frame_left.grid_rowconfigure(8, minsize=20)
frame_left.grid_rowconfigure(11, minsize=10)
"""frame section"""

tv_lbl_info1 = StringVar(value='ปีที่เลือก(จังหวัด)')
"""frame info section"""
# grid layout (1x1) info menu lbl
lbl_info_1 = CTkLabel(frame_info, textvariable=tv_lbl_info1, height=38, corner_radius=6,
                      text_font=myfont(18),
                      fg_color=("#ffffff", "gray38"),
                      justify=CENTER)
lbl_info_1.grid(row=0, column=0, sticky="nsew", padx=10, pady=10, ipady=1)
"""frame info section"""

"""#################frame graph section#################"""
tv_frame_right_m0 = StringVar(value='menu')
tv_frame_right_m1 = StringVar(value='menu')
tv_frame_right_m2 = StringVar(value='menu')
tv_frame_right_v0 = StringVar(value='0')
tv_frame_right_v1 = StringVar(value='0')
tv_frame_right_v2 = StringVar(value='0')

# three layout graph menu
frame_single = CTkFrame(frame_right)
frame_compare = CTkFrame(frame_right)
frame_overall = CTkFrame(frame_right)

# single mode variable
tv_lbl_path_single = StringVar(value='ยังไม่ได้เลือกไฟล์ข้อมูล')
tv_product_single = StringVar(value='0')
tv_proportion_single = StringVar(value='0')
tv_area_single = StringVar(value='0')
tv_yield_single = StringVar(value='0')
province_lst_single = ['เลือกไฟล์ก่อน']
data_single = {}

# compare mode variable
data_comp1 = {}
data_comp2 = {}
year_lst_comp = []
province_lst_comp = ['เลือกไฟล์ก่อน']

# overall mode variable
tv_lbl_path_ova = StringVar(value='ยังไม่ได้เลือกไฟล์ข้อมูล')
tv_sum_product = StringVar(value='0')
tv_sum_area = StringVar(value='0')
tv_mean_yield = StringVar(value='0')
province_lst_ova = ['เลือกไฟล์ก่อน']
province_graph = []
product_ova = []
product_ova2 = []
proportion_ova = []
area_ova = []
yield_ova = []

lbl_province_sl = CTkLabel(frame_right, text="เลือกจังหวัด:", text_font=myfont(16))
lbl_province_sl.grid(row=0, column=2, columnspan=1, pady=20, padx=10, sticky="nwe")
combo_province_single = CTkComboBox(frame_right, values=province_lst_single, text_font=myfont(12))
combo_province_comp = CTkComboBox(frame_right, values=province_lst_comp, text_font=myfont(12))

btn_download = CTkButton(frame_right, image=img_download, compound='left',
                         text="บันทึกกราฟ", text_font=myfont(16))
btn_download.grid(row=2, column=2, pady=10, padx=20, sticky="nwe")

btn_change_graph = CTkButton(frame_right, image=img_refresh, compound='left',
                             text="เปลี่ยนกราฟ", text_font=myfont(16))
btn_change_graph.grid(row=3, column=2, pady=10, padx=20, sticky="nwe")

"""Data visualization section"""


def mode_single():
    global frame_single, frame_compare, frame_overall
    global lbl_province_sl, combo_province_single, combo_province_comp
    global btn_change_graph
    frame_compare.grid_forget()
    frame_overall.grid_forget()
    btn_change_graph.grid_forget()
    combo_province_comp.grid_forget()
    frame_single.grid(row=1, column=0, columnspan=2, rowspan=7, pady=5, padx=5, sticky="nsew")
    lbl_province_sl.grid(row=0, column=2, columnspan=1, pady=20, padx=10, sticky="nwe")
    combo_province_single.grid(row=1, column=2, pady=10, padx=20, sticky="nwe")
    tv_lbl_info1.set('ปีที่เลือก(จังหวัด)')
    config['DISPLAYMODE'] = {'mode': '0'}
    with open(os.path.join(program_path, 'config\\cfg.ini'), 'w') as f:
        config.write(f)


def mode_compare():
    global frame_single, frame_compare, frame_overall
    global lbl_province_sl, combo_province_single, combo_province_comp
    global btn_change_graph
    frame_single.grid_forget()
    frame_overall.grid_forget()
    combo_province_single.grid_forget()
    btn_change_graph.grid(row=3, column=2, pady=10, padx=20, sticky="nwe")
    btn_change_graph.configure(command=lambda: update_chart_comp(None, change=True))
    frame_compare.grid(row=1, column=0, columnspan=2, rowspan=7, pady=5, padx=5, sticky="nsew")
    lbl_province_sl.grid(row=0, column=2, columnspan=1, pady=20, padx=10, sticky="nwe")
    combo_province_comp.grid(row=1, column=2, pady=10, padx=20, sticky="nwe")
    tv_lbl_info1.set('เปรียบเทียบปี(จังหวัด)')
    config['DISPLAYMODE'] = {'mode': '1'}
    with open(os.path.join(program_path, 'config\\cfg.ini'), 'w') as f:
        config.write(f)


def mode_overall():
    global frame_single, frame_compare, frame_overall
    global lbl_province_sl, combo_province_single, combo_province_comp
    global btn_change_graph
    frame_single.grid_forget()
    frame_compare.grid_forget()
    lbl_province_sl.grid_forget()
    combo_province_single.grid_forget()
    combo_province_comp.grid_forget()
    btn_change_graph.grid(row=3, column=2, pady=10, padx=20, sticky="nwe")
    btn_change_graph.configure(command=lambda: update_chart_overall())
    frame_overall.grid(row=1, column=0, columnspan=2, rowspan=7, pady=5, padx=5, sticky="nsew")
    tv_lbl_info1.set('ปีที่เลือก(ทั้งประเทศ)')
    config['DISPLAYMODE'] = {'mode': '2'}
    with open(os.path.join(program_path, 'config\\cfg.ini'), 'w') as f:
        config.write(f)


# developer info
def about_me():
    msg.showinfo('2022 KMITL. Programming Fundamental Project.',
                 'Name: Wiraphat Prasomphong\n' +
                 'Student ID: 65015143\n' +
                 'Computer Engineering DT.\n' +
                 "King Mongkut's Institute of Technology Ladkrabang."
                 )


# open the setting menu
def setting():
    setting = CTk()
    setting.title('ตั้งค่าโปรแกรม')
    setting.iconbitmap(os.path.join(program_path, 'img\\icon.ico'))
    # setting.attributes("-topmost", True)
    setting.geometry(center_screen(setting, 400, 580))

    # frame
    st_frame = CTkFrame(setting)
    st_frame.pack(pady=20, padx=60, fill='both', expand=True)

    st_title_lbl = CTkLabel(st_frame, text='ตั้งค่าการแสดงผล', justify=CENTER, text_font=myfont(16))
    st_title_lbl.pack(pady=12, padx=10)

    tv_st_map = StringVar(value=config.get('SETTING', 'map'))
    st_map_lbl = CTkLabel(st_frame, text='โหมดการแสดงผลของแผนที่', justify=CENTER, text_font=myfont(12))
    st_map_lbl.pack(pady=12, padx=10)
    st_map_0 = CTkRadioButton(st_frame, text='ไม่แสดงแผนที่ ', variable=tv_st_map, value='0', text_font=myfont(12))
    st_map_0.pack(pady=5, padx=10)
    st_map_1 = CTkRadioButton(st_frame, text='แผนที่แบบปกติ', variable=tv_st_map, value='1', text_font=myfont(12))
    st_map_1.pack(pady=5, padx=10)
    st_map_2 = CTkRadioButton(st_frame, text='แผนที่ดาวเทียม', variable=tv_st_map, value='2', text_font=myfont(12))
    st_map_2.pack(pady=5, padx=10)

    st_graphbg_default = config.get('SETTING', 'graphbg')
    st_graphbg = ['สีขาว', 'สีเทาอ่อน']
    st_graphbg_lbl = CTkLabel(st_frame, text='สีพื้นหลังของกราฟ', justify=CENTER, text_font=myfont(16))
    st_graphbg_lbl.pack(pady=12, padx=10)
    st_graph_combo = CTkComboBox(st_frame, values=st_graphbg, text_font=myfont(12))
    if st_graphbg_default == '0':
        st_graph_combo.set('สีขาว')
    else:
        st_graph_combo.set('สีเทาอ่อน')
    st_graph_combo.pack(pady=5, padx=10)

    st_color_set_default = config.get('SETTING', 'colorset')
    st_color = ['สีแบบไล่ระดับ1', 'สีแบบไล่ระดับ2',
                'สีแบบไล่ระดับ3', 'สีแบบไล่ระดับ4',
                'สีแบบค่าเริ่มต้น']
    st_color_lbl = CTkLabel(st_frame, text='ธีมสีของกราฟ', justify=CENTER, text_font=myfont(16))
    st_color_lbl.pack(pady=12, padx=10)
    st_color_combo = CTkComboBox(st_frame, values=st_color, text_font=myfont(12))
    if st_color_set_default == '0':
        st_color_combo.set('สีแบบไล่ระดับ1')
    elif st_color_set_default == '1':
        st_color_combo.set('สีแบบไล่ระดับ2')
    elif st_color_set_default == '2':
        st_color_combo.set('สีแบบไล่ระดับ3')
    elif st_color_set_default == '3':
        st_color_combo.set('สีแบบไล่ระดับ4')
    else:
        st_color_combo.set('สีแบบค่าเริ่มต้น')
    st_color_combo.pack(pady=5, padx=10)

    def save_setting():
        config['SETTING']['map'] = tv_st_map.get()
        config['SETTING']['graphbg'] = str(st_graphbg.index(st_graph_combo.get()))
        config['SETTING']['colorset'] = str(st_color.index(st_color_combo.get()))
        with open(os.path.join(program_path, 'config\\cfg.ini'), 'w') as f:
            config.write(f)
        msg.showinfo('บันทึกการตั้งค่า', 'บันทึกการตั้งค่าเรียบร้อยแล้ว โปรแกรมจะปิดตัวเองเพื่อให้การตั้งค่ามีผล')
        on_closing()

    st_save_btn = CTkButton(st_frame, text='บันทึกการตั้งค่า', text_font=myfont(18), command=save_setting)
    st_save_btn.pack(pady=12, padx=10)

    setting.mainloop()


# on startup mode change
startup_mode = {'0': mode_single, '1': mode_compare, '2': mode_overall}
startup_mode[config.get('DISPLAYMODE', 'mode')]()
# on startup mode change
"""Data visualization section"""

"""sigle mode"""
frame_single.columnconfigure(1, weight=1)


def select_file():
    global province_lst_single
    global combo_province_single
    global data_single
    file_path = filedialog.askopenfilename(filetypes=[('Excel', 'data_*.xlsx')])
    if file_path:
        try:
            tv_lbl_path_single.set(file_path)
            excel = pd.read_excel(file_path, skiprows=2)
            # data get from excel
            province = excel['จังหวัด'].values.tolist()
            product_single = excel['ผลผลิต(ตัน)'].values.tolist()
            proportion_single = excel['สัดส่วน'].values.tolist()
            area_single = excel['เนื้อที่เก็บเกี่ยว(ไร่)'].values.tolist()
            yield_single = excel['ผลผลิตต่อเนื้อที่เก็บเกี่ยว(กก.)'].values.tolist()
            # add data to dict
            for i in range(len(province)):
                data_single[province[i]] = {'ผลผลิต(ตัน)': product_single[i],
                                            'สัดส่วน': proportion_single[i] * 100,
                                            'เนื้อที่เก็บเกี่ยว(ไร่)': area_single[i],
                                            'ผลผลิตต่อเนื้อที่เก็บเกี่ยว(กก.)': yield_single[i]
                                            }

            year = os.path.basename(file_path)
            year = year.replace('.xlsx', '')
            year = year.replace('data_', '')
            province_lst_single = province
            combo_province_single.configure(values=province_lst_single, command=update_single)
            combo_province_single.set(province_lst_single[0])
            btn_download.configure(command=lambda: save_figure(pie_single))
            update_single(None)
        except Exception as e:
            msg.showerror('เกิดข้อผิดพลาด', f'เกิดข้อผิดพลาด {e}')


frame_file_select = CTkFrame(frame_single)
frame_file_select.grid(row=1, column=0, columnspan=2, padx=10, pady=10)
lbl_frame_file_select = CTkLabel(frame_file_select, textvariable=tv_lbl_path_single, text_font=myfont(12))
lbl_frame_file_select.grid(row=0, column=0, padx=10, pady=10)
btn_frame_file_select = CTkButton(frame_file_select, text="เลือกไฟล์", text_font=myfont(12), command=select_file,
                                  image=img_document, compound='left')
btn_frame_file_select.grid(row=0, column=1, padx=10, pady=10)

lbl_product = CTkLabel(frame_single, text='ผลผลิต(ตัน)', text_font=myfont(26))
lbl_product.grid(row=2, column=0, sticky="w", padx=10, pady=10)
lbl_area = CTkLabel(frame_single, text='เนื้อที่เก็บเกี่ยว(ไร่)', text_font=myfont(26))
lbl_area.grid(row=3, column=0, sticky="w", padx=10, pady=10)
lbl_yield = CTkLabel(frame_single, text='ผลผลิตต่อเนื้อที่เก็บเกี่ยว(กก.)', text_font=myfont(26))
lbl_yield.grid(row=4, column=0, sticky="w", padx=10, pady=10)

lbl_product_value = CTkLabel(frame_single, textvariable=tv_product_single, text_font=myfont(26))
lbl_product_value.grid(row=2, column=1, sticky="e", padx=10, pady=10)
lbl_area_value = CTkLabel(frame_single, textvariable=tv_area_single, text_font=myfont(26))
lbl_area_value.grid(row=3, column=1, sticky="e", padx=10, pady=10)
lbl_yield_value = CTkLabel(frame_single, textvariable=tv_yield_single, text_font=myfont(26))
lbl_yield_value.grid(row=4, column=1, sticky="e", padx=10, pady=10)

pie_single = plt.figure(figsize=(5, 5), dpi=90)
pie_single.add_subplot(111).pie([10, 100], colors=color_dict['colorset'][config.get('SETTING', 'colorset')],
                                autopct='%0.2f%%', pctdistance=1.2, startangle=90)
pie_single.legend(['สัดส่วนของจังหวัด', 'สัดส่วนทั้งหมด'], loc='upper right', bbox_to_anchor=(0.9, 1.0))
pie_single.set_facecolor(color_dict['graphbg'][config.get('SETTING', 'graphbg')])
pie_single_figure = FigureCanvasTkAgg(pie_single, frame_single)
pie_single.canvas.mpl_connect('button_press_event', lambda event: save_figure(pie_single))

if config.get('SETTING', 'map') != '0':
    pie_single_figure.get_tk_widget().grid(row=5, column=0, rowspan=3, sticky="nsew", padx=10, pady=10)
    pie_single.canvas.draw()
    map_sat = "https://mt0.google.com/vt/lyrs=s&hl=en&x={x}&y={y}&z={z}&s=Ga"
    map_normal = "https://mt0.google.com/vt/lyrs=m&hl=en&x={x}&y={y}&z={z}&s=Ga"
    map_dict = {'1': map_normal, '2': map_sat}
    map_widget = tkmap.TkinterMapView(frame_single, corner_radius=0)
    map_widget.set_tile_server(map_dict[config.get('SETTING', 'map')], max_zoom=22)
    marker = map_widget.set_address(f'สิงห์บุรี, ประเทศไทย', marker=True, text='สิงห์บุรี')
    map_widget.grid(row=5, column=1, rowspan=3, sticky="nsew", padx=10, pady=10)
else:
    pie_single_figure.get_tk_widget().grid(row=5, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)
    pie_single.canvas.draw()


def update_single(event):
    tv_product_single.set(f"{data_single[combo_province_single.get()]['ผลผลิต(ตัน)']:,} ตัน")
    tv_proportion_single.set(f"{data_single[combo_province_single.get()]['สัดส่วน']:,} %")
    tv_area_single.set(f"{data_single[combo_province_single.get()]['เนื้อที่เก็บเกี่ยว(ไร่)']:,} ไร่")
    tv_yield_single.set(f"{data_single[combo_province_single.get()]['ผลผลิตต่อเนื้อที่เก็บเกี่ยว(กก.)']:,} กก./ไร่")
    pie_single.clear()
    pie_single.add_subplot(111).pie([data_single[combo_province_single.get()]['สัดส่วน'], 100],
                                    colors=color_dict['colorset'][config.get('SETTING', 'colorset')], autopct='%0.2f%%',
                                    pctdistance=1.2, startangle=90)
    pie_single.legend([f'สัดส่วนผลผลิตของจังหวัด{combo_province_single.get()}', 'สัดส่วนผลผลิตทั้งหมด'],
                      loc='upper right', bbox_to_anchor=(0.9, 1.0))
    pie_single.canvas.draw()

    if connected():
        if config.get('SETTING', 'map') != '0':
            map_widget.set_address(f'{combo_province_single.get()}, ประเทศไทย', marker=True,
                                   text=combo_province_single.get())
    else:
        msg.showwarning('ไม่พบการเชื่อมต่ออินเทอร์เน็ต', 'ไม่สามารถอัพเดทแผนที่ได้ โปรดตรวจสอบการเชื่อมต่ออินเทอร์เน็ต')


"""single mode"""

"""compare mode"""
frame_compare.columnconfigure(1, weight=1)
path_comp1, path_comp2 = '', ''


def select_file_comp(is_path1):
    global path_comp1, path_comp2
    file_path = filedialog.askopenfilename(filetypes=[('Excel', 'data_*.xlsx')])
    if file_path:
        if is_path1:
            path_comp1 = file_path
            tv_lbl_path_comp1.set(file_path)
        else:
            path_comp2 = file_path
            tv_lbl_path_comp2.set(file_path)


def update_comp():
    global data_comp1, data_comp2
    global year_lst_comp, province_lst_comp
    if path_comp1 != '' and path_comp2 != '':
        if path_comp1 != path_comp2:
            year_lst_comp = []
            excel1 = pd.read_excel(path_comp1, skiprows=2)
            excel2 = pd.read_excel(path_comp2, skiprows=2)

            product1 = excel1['ผลผลิต(ตัน)'].values.tolist()
            product2 = excel2['ผลผลิต(ตัน)'].values.tolist()
            area1 = excel1['เนื้อที่เก็บเกี่ยว(ไร่)'].values.tolist()
            area2 = excel2['เนื้อที่เก็บเกี่ยว(ไร่)'].values.tolist()
            yield1 = excel1['ผลผลิตต่อเนื้อที่เก็บเกี่ยว(กก.)'].values.tolist()
            yield2 = excel2['ผลผลิตต่อเนื้อที่เก็บเกี่ยว(กก.)'].values.tolist()
            proportion1 = excel1['สัดส่วน'].values.tolist()
            proportion2 = excel2['สัดส่วน'].values.tolist()
            prov1 = excel1['จังหวัด'].values.tolist()
            prov2 = excel2['จังหวัด'].values.tolist()
            province = min(len(prov1), len(prov2))
            province = prov1[:province] if province == len(prov1) else prov2[:province]
            province_lst_comp = province

            combo_province_comp.configure(values=province_lst_comp, command=update_chart_comp)
            combo_province_comp.set(province_lst_comp[0])
            btn_download.configure(command=lambda: save_figure(chart_comp))
            btn_change_graph.configure(command=lambda: update_chart_comp(None, change=True))

            year1 = os.path.basename(path_comp1).replace('data_', '').replace('.xlsx', '')
            year2 = os.path.basename(path_comp2).replace('data_', '').replace('.xlsx', '')
            year_lst_comp.append(year1)
            year_lst_comp.append(year2)

            data_comp1 = {}
            data_comp2 = {}
            for i in range(len(province)):
                data_comp1[province[i]] = {'ผลผลิต(ตัน)': product1[i],
                                           'สัดส่วน': proportion1[i] * 100,
                                           'เนื้อที่เก็บเกี่ยว(ไร่)': area1[i],
                                           'ผลผลิตต่อเนื้อที่เก็บเกี่ยว(กก.)': yield1[i]
                                           }
                data_comp2[province[i]] = {'ผลผลิต(ตัน)': product2[i],
                                           'สัดส่วน': proportion2[i] * 100,
                                           'เนื้อที่เก็บเกี่ยว(ไร่)': area2[i],
                                           'ผลผลิตต่อเนื้อที่เก็บเกี่ยว(กก.)': yield2[i]
                                           }

            update_chart_comp(None)

        else:
            msg.showwarning('เลือกไฟล์ซ้ำกัน', 'โปรดเลือกไฟล์ที่แตกต่างกัน')
    else:
        msg.showwarning('ไม่พบไฟล์', 'โปรดเลือกไฟล์ที่ต้องการเปรียบเทียบ')


tv_lbl_path_comp1 = StringVar(value='ไม่ได้เลือกไฟล์')
tv_lbl_path_comp2 = StringVar(value='ไม่ได้เลือกไฟล์')

frame_file_select_comp = CTkFrame(frame_compare)
frame_file_select_comp.grid(row=1, rowspan=1, column=0, columnspan=2, padx=10, pady=10)
lbl_frame_file_select_comp1 = CTkLabel(frame_file_select_comp,
                                       textvariable=tv_lbl_path_comp1, text_font=myfont(12))
lbl_frame_file_select_comp1.grid(row=0, column=0, padx=10, pady=10)
btn_frame_file_select_comp1 = CTkButton(frame_file_select_comp, text="เลือกไฟล์", text_font=myfont(12),
                                        command=lambda: select_file_comp(True), image=img_document, compound='left')
btn_frame_file_select_comp1.grid(row=0, column=1, padx=10, pady=10)
lbl_frame_file_select_comp2 = CTkLabel(frame_file_select_comp, textvariable=tv_lbl_path_comp2, text_font=myfont(12))
lbl_frame_file_select_comp2.grid(row=1, column=0, padx=10, pady=10)
btn_frame_file_select_comp2 = CTkButton(frame_file_select_comp, text="เลือกไฟล์", text_font=myfont(12),
                                        command=lambda: select_file_comp(False), image=img_document, compound='left')
btn_frame_file_select_comp2.grid(row=1, column=1, padx=10, pady=10)
btn_confirm_comp = CTkButton(frame_compare, text="อัพเดท", text_font=myfont(12),
                             command=update_comp, image=img_comp, compound='left')
btn_confirm_comp.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

fig_comp = plt.Figure(figsize=(6, 6), dpi=100)
chart_comp = fig_comp.add_subplot(111)
chart_comp.barh(['2001', '2544'], [15, 27], height=0.25,
                color=color_dict['colorset'][config.get('SETTING', 'colorset')])
chart_comp.set_title('กราฟตัวอย่าง')
chart_comp_figure = FigureCanvasTkAgg(fig_comp, frame_compare)
chart_comp_figure.get_tk_widget().grid(row=5, column=0, rowspan=3, columnspan=2, sticky="nsew", padx=10, pady=10)
fig_comp.set_facecolor(color_dict['graphbg'][config.get('SETTING', 'graphbg')])
fig_comp.canvas.mpl_connect('button_press_event', lambda event: save_figure(fig_comp))
fig_comp.canvas.draw()

update_comp_graph = 0


def update_chart_comp(event, change=False):
    try:
        global fig_comp, chart_comp, update_comp_graph
        chart_comp.clear()
        if change:
            update_comp_graph += 1
        if update_comp_graph == 0:
            fig_comp.suptitle(
                f'ผลผลิต(ตัน)ระหว่างปี {year_lst_comp[0]} และปี {year_lst_comp[1]}\nของจังหวัด{combo_province_comp.get()}')
            rect = chart_comp.barh(year_lst_comp, [data_comp1[combo_province_comp.get()]['ผลผลิต(ตัน)'],
                                                   data_comp2[combo_province_comp.get()]['ผลผลิต(ตัน)']], height=0.25,
                                   color=color_dict['colorset'][config.get('SETTING', 'colorset')])
            labels = [data_comp1[combo_province_comp.get()]['ผลผลิต(ตัน)'],
                      data_comp2[combo_province_comp.get()]['ผลผลิต(ตัน)']]
            chart_comp.bar_label(rect, labels=labels, label_type='center')
            fig_comp.canvas.draw()
        elif update_comp_graph == 1:
            fig_comp.suptitle(
                f'เนื้อที่เก็บเกี่ยว(ไร่)\nปี {year_lst_comp[0]} และปี {year_lst_comp[1]} จังหวัด{combo_province_comp.get()}')
            rect = chart_comp.bar(year_lst_comp, [data_comp1[combo_province_comp.get()]['เนื้อที่เก็บเกี่ยว(ไร่)'],
                                                  data_comp2[combo_province_comp.get()]['เนื้อที่เก็บเกี่ยว(ไร่)']],
                                  width=0.25,
                                  color=color_dict['colorset'][config.get('SETTING', 'colorset')])
            labels = [data_comp1[combo_province_comp.get()]['เนื้อที่เก็บเกี่ยว(ไร่)'],
                      data_comp2[combo_province_comp.get()]['เนื้อที่เก็บเกี่ยว(ไร่)']]
            chart_comp.bar_label(rect, labels=labels, label_type='center')
            fig_comp.canvas.draw()
        elif update_comp_graph == 2:
            fig_comp.suptitle(
                f'ผลผลิตต่อเนื้อที่เก็บเกี่ยว(กก.)\nปี {year_lst_comp[0]} และปี {year_lst_comp[1]} จังหวัด{combo_province_comp.get()}')
            rect = chart_comp.bar(year_lst_comp,
                                  [data_comp1[combo_province_comp.get()]['ผลผลิตต่อเนื้อที่เก็บเกี่ยว(กก.)'],
                                   data_comp2[combo_province_comp.get()]['ผลผลิตต่อเนื้อที่เก็บเกี่ยว(กก.)']],
                                  width=0.25,
                                  color=color_dict['colorset'][config.get('SETTING', 'colorset')])
            labels = [data_comp1[combo_province_comp.get()]['ผลผลิตต่อเนื้อที่เก็บเกี่ยว(กก.)'],
                      data_comp2[combo_province_comp.get()]['ผลผลิตต่อเนื้อที่เก็บเกี่ยว(กก.)']]
            chart_comp.bar_label(rect, labels=labels, label_type='center')
            fig_comp.canvas.draw()
            if change:
                update_comp_graph = -1
    except IndexError as e:
        pass


"""compare mode"""

"""overall mode"""
frame_overall.columnconfigure(1, weight=1)


def select_file_overall():
    global province_lst_ova
    # global combo_province
    global btn_change_graph, btn_download
    global product_ova, proportion_ova, area_ova, yield_ova, product_ova2
    global province_graph
    file_path = filedialog.askopenfilename(filetypes=[('Excel', 'data_*.xlsx')])
    if file_path:
        try:
            tv_lbl_path_ova.set(file_path)
            excel = pd.read_excel(file_path, skiprows=2)
            # data get from excel
            province = excel['จังหวัด'].values.tolist()
            # product_ova =
            product = excel['ผลผลิต(ตัน)'].values.tolist()
            area_ova = excel['เนื้อที่เก็บเกี่ยว(ไร่)'].values.tolist()
            yield_ova = excel['ผลผลิตต่อเนื้อที่เก็บเกี่ยว(กก.)'].values.tolist()

            proportion_ova = []
            province_graph = []
            for i in excel['สัดส่วน'].values.tolist():
                if float(i) * 100 > 3:
                    proportion_ova.append(i)

            for i in range(len(proportion_ova)):
                province_graph.append(province[i])

            # product_ova = product
            product_ova = []
            for i in range(10):
                product_ova.append(product[i])
            product_ova.append(sum(product) - sum(product_ova))
            product_ova2 = product

            proportion_ova.append(sum(excel['สัดส่วน'].values.tolist()) - sum(proportion_ova))
            province_graph.append('จังหวัดที่ต่ำกว่า 3%')

            year = os.path.basename(file_path)
            year = year.replace('.xlsx', '')
            year = year.replace('data_', '')
            province_lst_ova = province
            btn_change_graph.configure(command=lambda: update_chart_overall())
            btn_download.configure(command=lambda: save_figure(fig_ova))
            update_overall(None)
        except Exception as e:
            msg.showerror('เกิดข้อผิดพลาด', f'เกิดข้อผิดพลาด {e}')


frame_file_select_overall = CTkFrame(frame_overall)
frame_file_select_overall.grid(row=1, column=0, columnspan=2, padx=10, pady=10)
lbl_frame_file_select_overall = CTkLabel(frame_file_select_overall, textvariable=tv_lbl_path_ova, text_font=myfont(12))
lbl_frame_file_select_overall.grid(row=0, column=0, padx=10, pady=10)
btn_frame_file_select_overall = CTkButton(frame_file_select_overall, text="เลือกไฟล์", text_font=myfont(12),
                                          command=select_file_overall, image=img_document, compound='left')
btn_frame_file_select_overall.grid(row=0, column=1, padx=10, pady=10)

lbl_sum_product = CTkLabel(frame_overall, text='ผลผลิตรวม(ล้านตัน)', text_font=myfont(26))
lbl_sum_product.grid(row=2, column=0, sticky="w", padx=10, pady=10)
lbl_sum_area = CTkLabel(frame_overall, text='เนื้อที่เก็บเกี่ยวรวม(ไร่)', text_font=myfont(26))
lbl_sum_area.grid(row=3, column=0, sticky="w", padx=10, pady=10)
lbl_mean_yield = CTkLabel(frame_overall, text='ค่าเฉลี่ยผลผลิตจากเนื้อที่เก็บเกี่ยวทั้งหมด(กก.)', text_font=myfont(26))
lbl_mean_yield.grid(row=4, column=0, sticky="w", padx=10, pady=10)

lbl_sum_product_value = CTkLabel(frame_overall, textvariable=tv_sum_product, text_font=myfont(26))
lbl_sum_product_value.grid(row=2, column=1, sticky="e", padx=10, pady=10)
lbl_sum_area_value = CTkLabel(frame_overall, textvariable=tv_sum_area, text_font=myfont(26))
lbl_sum_area_value.grid(row=3, column=1, sticky="e", padx=10, pady=10)
lbl_mean_yield_value = CTkLabel(frame_overall, textvariable=tv_mean_yield, text_font=myfont(26))
lbl_mean_yield_value.grid(row=4, column=1, sticky="e", padx=10, pady=10)

fig_ova = plt.figure(figsize=(5, 5), dpi=100)
fig_ova.suptitle('กราฟตัวอย่าง', fontsize=16)
chart_ova = fig_ova.add_subplot(111)
chart_ova.pie([10, 100], autopct='%1.1f%%', pctdistance=1.2, startangle=90,
              colors=color_dict['colorset'][config.get('SETTING', 'colorset')])
chart_ova_figure = FigureCanvasTkAgg(fig_ova, frame_overall)
chart_ova_figure.get_tk_widget().grid(row=5, column=0, rowspan=3, columnspan=2, sticky="nsew", padx=10, pady=10)
fig_ova.set_facecolor(color_dict['graphbg'][config.get('SETTING', 'graphbg')])
fig_ova.canvas.mpl_connect('button_press_event', lambda event: save_figure(fig_ova))
fig_ova.canvas.draw()

tha = pd.read_excel(os.path.join(program_path, 'config\\thaidata.xlsx'))
sec = tha['ภาค'].values.tolist()
prov = tha['จังหวัด'].values.tolist()
sec_dict = dict(zip(prov, sec))
sector = tha['ภาค'].unique().tolist()
sector_ova = [0, 0, 0, 0, 0, 0]

update_ova = 0


def update_chart_overall():
    try:
        if tv_lbl_path_ova.get() == 'ยังไม่ได้เลือกไฟล์ข้อมูล' or tv_lbl_path_ova.get() == '':
            return
        global update_ova, fig_ova
        chart_ova.clear()
        if update_ova == 0:  # pie proportion
            fig_ova.suptitle('สัดส่วนผลผลิตแต่ละจังหวัด', fontsize=16)
            chart_ova.pie(proportion_ova, autopct='%1.1f%%', pctdistance=1.2, startangle=90,
                          explode=[0.1] * len(proportion_ova),
                          colors=color_dict['colorset'][config.get('SETTING', 'colorset')], normalize=True)
            chart_ova.legend(province_graph, loc='upper right', ncol=2, bbox_to_anchor=(0.0005, 1.15))
            fig_ova.canvas.draw()
        elif update_ova == 1:  # pie product
            legend_lbl = []
            for i in range(len(product_ova) - 1):
                legend_lbl.append(f'{province_lst_ova[i]} {product_ova[i]:,} ตัน')
            legend_lbl.append(f'จังหวัดอื่นๆ {product_ova[-1]:,} ตัน')
            pie_ova_labels = province_lst_ova[:10] + ['จังหวัดอื่นๆ']
            fig_ova.suptitle('ผลผลิตแต่ละจังหวัด(ตัน)', fontsize=16)
            chart_ova.pie(product_ova, labels=pie_ova_labels, startangle=90, normalize=True, explode=[
                                                                                                         0.1] * len(
                product_ova), colors=color_dict['colorset'][config.get('SETTING', 'colorset')])
            chart_ova.legend(legend_lbl, loc='upper left', bbox_to_anchor=(1.2, 1.15))
            fig_ova.canvas.draw()
        else:
            for i in range(len(product_ova2)):
                if sec_dict.get(province_lst_ova[i]) == 'กรุงเทพและปริมณฑล':
                    sector_ova[0] += product_ova2[i]
                elif sec_dict.get(province_lst_ova[i]) == 'ภาคกลาง':
                    sector_ova[1] += product_ova2[i]
                elif sec_dict.get(province_lst_ova[i]) == 'ภาคเหนือ':
                    sector_ova[2] += product_ova2[i]
                elif sec_dict.get(province_lst_ova[i]) == 'ภาคตะวันออกเฉียงเหนือ':
                    sector_ova[3] += product_ova2[i]
                elif sec_dict.get(province_lst_ova[i]) == 'ภาคใต้':
                    sector_ova[4] += product_ova2[i]
                elif sec_dict.get(province_lst_ova[i]) == 'ภาคตะวันออก':
                    sector_ova[5] += product_ova2[i]

            fig_ova.suptitle('ผลผลิตแบ่งตามภาค', fontsize=16)
            chart_ova.pie(sector_ova, autopct='%1.1f%%', pctdistance=1.2, startangle=90,
                          explode=[0.1] * len(sector_ova),
                          colors=color_dict['colorset'][config.get('SETTING', 'colorset')],
                          normalize=True)
            chart_ova.legend(sector, loc='upper left', bbox_to_anchor=(1.2, 1.15))
            fig_ova.canvas.draw()
            update_ova = -1
        update_ova += 1
    except IndexError as e:
        pass


def update_overall(event):
    tv_sum_product.set(f"{round(sum(product_ova) / 1000000, 2):,} ล้านตัน")
    tv_sum_area.set(f"{sum(area_ova):,} ไร่")
    tv_mean_yield.set(f"{round(sum(yield_ova) / len(yield_ova), 2):,} กก./ไร่")
    update_chart_overall()


"""overall mode"""


def save_figure(fig):
    global chart_ova
    global pie_single
    file = filedialog.asksaveasfile(mode='w', defaultextension=".png",
                                    filetypes=(("PNG file", "*.png"), ("All Files", "*.*")))
    if file:
        fig.savefig(file.name)


"""#################frame graph section#################"""

"""menu section"""
lbl_menu = CTkLabel(frame_left, text="เมนูการแสดงผล", text_font=myfont(18))
lbl_menu.grid(row=1, column=0, pady=10, padx=10)
# button for select mode
btn_1 = CTkButton(frame_left, text="ปีที่เลือก(จังหวัด)", image=img_chart_pie, compound="right",
                  text_font=myfont(), command=mode_single)
btn_1.grid(row=2, column=0, pady=10, padx=20, sticky="we")
btn_2 = CTkButton(frame_left, text="เปรียบเทียบปี(จังหวัด)", image=img_stats, compound="right",
                  text_font=myfont(), command=mode_compare)
btn_2.grid(row=3, column=0, pady=10, padx=20, sticky="we")
btn_3 = CTkButton(frame_left, text="ปีที่เลือก(ทั้งประเทศ)", image=img_chart_histogram, compound="right",
                  text_font=myfont(), command=mode_overall)
btn_3.grid(row=4, column=0, pady=10, padx=20, sticky="we")
btn_4 = CTkButton(frame_left, text="เกี่ยวกับผู้พัฒนา", image=img_user, compound="right",
                  text_font=myfont(), command=about_me)
btn_4.grid(row=6, column=0, pady=10, padx=20, sticky="we")
btn_5 = CTkButton(frame_left, text="ตั้งค่า", image=img_settings_image, compound="right",
                  text_font=myfont(), command=setting)
btn_5.grid(row=7, column=0, pady=10, padx=20, sticky="we")
# theme lbl
lbl_mode = CTkLabel(frame_left, text="โหมดการแสดงผล:", text_font=myfont(12))
lbl_mode.grid(row=9, column=0, pady=0, padx=20, sticky="we")
# change theme combobox
optionmenu_1 = CTkOptionMenu(frame_left, values=theme, command=change_theme, text_font=myfont())
optionmenu_1.grid(row=10, column=0, pady=10, padx=20, ipady=2, sticky="we")
""""menu section"""

root.mainloop()  # run mainloop of customtkinter
