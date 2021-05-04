
import xlsxwriter
from tempfile import TemporaryFile
import time
from moviepy.editor import *
from moviepy.video.io.ffmpeg_tools import ffmpeg_extract_subclip
from tkinter import *
import tkinter.font as font


def gui_multiple_box():
    def gui_box(title, user_question):
        master = Tk()
        master.title(title)
        e = Entry(master, width=10, font=('Helvetica', 50))
        e.pack(padx=100, pady=50)

        # define font
        myFont = font.Font(family='Helvetica', size=25, weight='bold')

        # myFont = font.Font(size=20)
        def callback():
            show_input = "File Name: {}".format(e.get())
            myLabel = Label(master, text=show_input)
            myLabel.pack(padx=10, pady=10)
            global text
            text = e.get()
            # print(e.get())  # This is the text you may want to use later
            e.delete(0, 'end')
            master.destroy()

        b = Button(master, text=user_question, bg='#0052cc', fg='#ffffff', height=5, width=30, command=callback)
        b['font'] = myFont
        b.pack(padx=10, pady=10)
        mainloop()

        return text

    check_camera_name = gui_box('file_name', 'Enter Camera Name \n(to select all type *)  ')
    print('camera_name = ', check_camera_name)
    check_date = gui_box('date', 'Enter date YYYY-MM-DD \n  (to select all type *)  ')
    print('date = ', check_date)
    check_time = gui_box('time', 'Enter Approximate Time HH.MM \n(to select all type *)  ')
    print('time = ', check_time)
    # check_length = gui_box('Video Length', 'Enter Length \n(to select all type *)  ')
    # print('video length = ', check_time)
    check_length = '*'

    return check_camera_name, check_date, check_time, check_length


def video_duration(file):
    return VideoFileClip(file).duration


def convert_time_to_min(time):
    hour = time.split('.')[0]
    minute = time.split('.')[1]
    converted_time = int(hour) * 60 + int(minute)
    return converted_time


def videos(FOLDER_PATH):
    for file in os.listdir(FOLDER_PATH):
        if file.endswith(".mp4") or file.endswith(".avi"):
            all_videos_list.append(file)
        # print(all_videos, sep="\n")
    return all_videos_list


def cameras(check_camera_name, FOLDER_PATH):
    for file_name in videos(FOLDER_PATH):
        if check_camera_name == file_name.split('+')[0]:
            # print(check_camera_name, sep="\n")
            camera_names_list.append(file_name)
        elif check_camera_name == '*':
            # print(file, sep="\n")
            camera_names_list.append(file_name)
    return camera_names_list  # Videos fulfilling the given condition of camera name


def date_fnc(check_date, check_camera_name, FOLDER_PATH):
    for file_name in cameras(check_camera_name, FOLDER_PATH):
        if check_date == file_name.split('+')[1]:
            # print(check_date, sep="\n")
            date_list.append(file_name)
        elif check_date == "*":
            # print(file, sep="\n")
            date_list.append(file_name)
    return date_list  # Videos fulfilling the condition of camera name and date


def time_fnc(check_time, check_date, check_camera_name, FOLDER_PATH):
    for file_name in date_fnc(check_date, check_camera_name, FOLDER_PATH):
        time = ".".join(file_name.split('+')[2].split('.', 2)[:2])
        # print(check_coverted_time)
        if check_time != "*":
            coverted_time = convert_time_to_min(time)
            if convert_time_to_min(check_time) in range(coverted_time - 30, coverted_time + 30):
                # print(check_minutes, sep="\n")
                minutes_list.append(file_name)
        elif check_time == "*":
            # print(file, sep="\n")
            minutes_list.append(file_name)
    return minutes_list  # Videos fulfilling condition of cam_nam, date and time


def length_fnc(check_length, check_time, check_date, check_camera_name, FOLDER_PATH):
    for file_name in time_fnc(check_time, check_date, check_camera_name, FOLDER_PATH):
        if check_length != "*":
            if video_duration(file_name) >= (int(end_time) - start_time):
                ffmpeg_extract_subclip(file_name, start_time, int(end_time), targetname=(file_name + ".mp4"))  # trims the video
                target_name = (file_name + ".mp4")
                length_list.append(target_name)
            else:
                print("The original video " + file_name + " duration is less than required clipping length")
        elif check_length == "*":
            length_list.append(file_name)  # appends these video in the list of required videos
    print(length_list, end='\n')
    if not length_list:
        print("THERE ARE NO VIDEOS MATCHING WITH YOUR DESCRIPTION")
        return length_list
    else:
        return length_list  # final list of videos which satisfy our conditions


if __name__ == "__main__":
    # FOLDER_PATH = r'E:\\indot videos 3\\65_2540+2019-03-07+21_to_0'
    FOLDER_PATH = r'E:\\indot videos'

    all_videos_list = []
    camera_names_list = []
    date_list = []
    minutes_list = []
    length_list = []
    file_names_list = []
    path_names_list = []
    # ------------ input variables ---------------#

    [check_camera_name, check_date, check_time, check_length] = gui_multiple_box()
    # [check_camera_name,check_date,check_time,check_length] = ["*", "*", "*", "*"]
    # check_camera_name = "65_2540"
    # check_date = "2019-10-03"
    # check_time = "6.59"
    # check_length = "5"

    # check_camera_name = input("ENTER CAMERA NAME : ")
    # check_date = input("ENTER DATE : ")
    # check_time = input("ENTER TIME : ")
    # check_length = input("ENTER LENGTH(in seconds) : ")
    # --------------------------------------------#
    start_time = 0
    end_time = check_length

    for fileName in length_fnc(check_length, check_time, check_date, check_camera_name, FOLDER_PATH):
        file_names_list.append(fileName)
        path_names_list.append(os.path.abspath(os.path.join(FOLDER_PATH, fileName)))

    # ---------------------- #
    # --MAKING EXCEL SHEET-- #
    # ---------------------- #

    # Create a new workbook and add a worksheet
    workbook = xlsxwriter.Workbook('requirement_satisfied_files.xlsx')
    worksheet = workbook.add_worksheet('Hyperlinks_and_names')

    # Add a format to use to highlight cells.
    bold_and_font = workbook.add_format({'bold': True, 'font_size': 24})
    big_font = workbook.add_format({'font_size': 14})
    hyperlink = workbook.add_format({'hyperlink': True, 'font_size': 14})

    worksheet.write(0, 0, "FILE NAME", bold_and_font)
    worksheet.write(0, 1, "FILE PATH", bold_and_font)
    worksheet.write(0, 2, "HYPERLINK", bold_and_font)

    for i, e in enumerate(file_names_list):
        worksheet.write(i+1, 0, e, big_font)
    for i, e in enumerate(path_names_list):
        worksheet.write(i+1, 1, e, big_font)
    for i, e in enumerate(path_names_list):
        worksheet.write(i+1, 2, "=HYPERLINK(B" + str(int(i)+2) + "," + "A" + str(int(i)+2) + ")", hyperlink)

    workbook.close()
    # excel_application_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
    # cmd = "open -a 'path' 'requirement_satisfied_files.xlsx'"
    cmd = 'requirement_satisfied_files.xlsx'
    time.sleep(1)
    os.system(cmd)
