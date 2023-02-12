from tkinter import messagebox
import socket
import configparser
import os

"""
                           __    _____    ___    __   _____   __   _  _     ____  
                          / /   | ____|  / _ \  /_ | | ____| /_ | | || |   |___ \ 
                         / /_   | |__   | | | |  | | | |__    | | | || |_    __) |
                        | '_ \  |___ \  | | | |  | | |___ \   | | |__   _|  |__ < 
                        | (_) |  ___) | | |_| |  | |  ___) |  | |    | |    ___) |
                         \___/  |____/   \___/   |_| |____/   |_|    |_|   |____/ 

    ###############################################################################################
    #                            ข้อจำกัดของการตรวจสอบ ERROR                                          #
    # รหัส     ความหมาย                                                                             #
    # 200     ปกติ                                                                                 #
    # 404     ไม่พบไฟล์บางส่วน สามารถซ่อมได้ เช่น cfg.ini หรือ ไฟล์ cfg.ini ไม่ถูกต้อง                          #
    # 500     ไม่พบไฟล์ ฟอนต์ รูปภาพ ไอคอน ของโปรแกรม หรือไฟล์ที่ทำให้ทำงานต่อไม่ได้ โปรแกรมจะปิดการทำงาน            #
    #                                                                                             #
    #                            ข้อจำกัดของการซ่อมไฟล์ ERROR                                          #                  
    # cfg.ini สามารถซ่อมได้                                                                          #                                
    #                            ตรวจสอบการเชื่อมต่อ internet                                          #
    # True    สามารถเชื่อมต่อได้                                                                        #
    # False   ไม่สามารถเชื่อมต่อได้                                                                      #
    ###############################################################################################

    .----------------.  .----------------. .----------------.  .----------------. .----------------. 
    | .--------------.|| .--------------. || .--------------.|| .--------------.||.--------------. |
    | |  ___  ____   ||||  ____    ____  ||||    _____      ||||  _________    ||||   _____      | |
    | | |_  ||_  _|  |||| |_   \  /   _| ||||    |_   _|    |||| |  _   _  |   ||||  |_   _|     | |
    | |   | |_/ /    ||||   |   \/   |   ||||      | |      |||| |_/ | | \_|   ||||    | |       | |
    | |   |  __'.    ||||   | |\  /| |   ||||      | |      ||||     | |       ||||    | |   _   | |
    | |  _| |  \ \_  ||||  _| |_\/_| |_  ||||     _| |_     ||||    _| |_      ||||   _| |__/ |  | |
    | | |____||____| |||| |_____||_____| ||||    |_____|    ||||   |_____|     ||||  |________|  | |
    | |              ||||                ||||               ||||               ||||              | |
    | '--------------'|| '--------------' ||'--------------' || '--------------'||'--------------' |
    '----------------'  '----------------' '----------------'  '----------------' '----------------' 
"""
# path of program
program_path = os.path.dirname(__file__)

# init config for read cfg.ini
config = configparser.ConfigParser(allow_no_value=True)

# variable for check error
theme = ['dark', 'light']
mode = ['0', '1', '2']
graph = ['0', '1']
color = ['0', '1', '2', '3', '4']
img = ['chart-histogram.png', 'chart-pie.png', 'compare.png', 'document.png',
       'download.png', 'icon.ico', 'refresh.png', 'settings.png', 'stats.png', 'user.png']
conf = ['cfg.ini', 'thaidata.xlsx', 'font\\NotoSerifThai.ttf']


# create new config file
def mkconf():
    with open(os.path.join(program_path, "config/cfg.ini"), 'w') as f:
        config['THEME'] = {'mode': 'dark'}
        config['DISPLAYMODE'] = {'mode': '0'}
        config['SETTING'] = {'colorset': '4', 'graphbg': '0', 'map': '1'}
        config.write(f)


# main checking file and folder error function
def check_error():
    err = {'err': {'200': 'normal state'}}
    # config dir
    if os.path.isdir(os.path.join(program_path, 'config')):
        for i in conf:
            if not os.path.isfile(os.path.join(program_path, 'config\\' + i)):
                if i == 'cfg.ini':
                    mkconf()
                    err['err'] = {'404': i}
                else:
                    err['err'] = {'500': i}
                break
    else:
        err['err'] = {'500': 'config directory'}
    # img dir
    if os.path.isdir(os.path.join(program_path, 'img')):
        for i in img:
            if not os.path.isfile(os.path.join(program_path, 'img\\' + i)):
                err['err'] = {'500': i}
                break
    else:
        err['err'] = {'500': 'img directory'}
    # check corupted config.ini
    if os.stat(os.path.join(program_path, 'config/cfg.ini')).st_size == 0:
        mkconf()
        err['err'] = {'404': 'ERR0 cfg.ini is corrupted'}
    else:
        config.read(os.path.join(program_path, 'config/cfg.ini'))
        if not config['THEME']['mode'] in theme or not config['DISPLAYMODE']['mode'] in mode or not config['SETTING']['colorset'] in color or \
                not config['SETTING']['graphbg'] in graph or not config['SETTING']['map'] in mode:
            mkconf()
            err['err'] = {'404': 'ERR1 cfg.ini is corrupted'}
    return err


# check internet connection
def connected():
    try:  # connect to the host -- tells us if the host is actually reachable
        socket.create_connection(("www.google.com", 80))
        return True  # connected
    except OSError:
        pass  # pass system error
    return False  # not connected


if __name__ == '__main__':
    messagebox.showerror('ERRCHK!', 'นี่ไม่ใช่โปรแกรมหลัก โปรดเปิดโปรแกรมด้วย main.py')
    #print(check_error())
    #print(connected())
