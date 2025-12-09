from pywinauto.application import Application
from pywinauto import mouse
from time import sleep
import win32api
import os
import csv
from datetime import datetime
from influxdb_client import Point, InfluxDBClient
from influxdb_client.client.write_api import SYNCHRONOUS
from dotenv import dotenv_values

# based on docs, I tried using Accessibility Insights for Windows
# and spyxx, but neither give info about the app (meaning it's mostly custom controls).
# as such, we need to mostly fall back to pixel math and mouse clicks rather
# than advanced processing pywinauto provides (except in certain cases)


# for debugging, print cursor position

def p():
    print(win32api.GetCursorPos())

# click the window at the location


def click(win, x, y, times=1):
    print(f"click({x}, {y}) times={times}")
    target = (x, y)
    win.restore()
    if times == 1:
        mouse.click(coords=target)
    else:
        mouse.double_click(coords=target)


def download_new_data():
    file_dir = None

    try:
        # spawn the app
        app = Application(backend='uia').start(
            "C:\\Program Files\\QUARTA-RAD\\RadexDC64\\RadexDC.exe")

        # retrive a handle to main window
        main_win = app['Please connect the device']

        # wait for radex dc to boot
        sleep(2)
        # main_win.wait('ready', timeout=15) # this never worked, idk why

        # select first device
        main_rect = main_win.rectangle()
        title_win = main_win.children()[0]
        title_rect = title_win.rectangle()

        click(main_win, main_rect.left + 40,
              main_rect.top + title_rect.height() + 40, 2)

        # wait for device options dialog to appear
        sleep(2)

        # click on download
        main_win = app.Dialog
        main_rect = main_win.rectangle()

        click(main_win, main_rect.left + 200, main_rect.top + 200)

        # wait for download dialog to process data from device
        sleep(10)

        # click on save
        click(main_win, main_rect.left + 200 - 100, main_rect.top + 200 + 70)

        # wait for save dialog to appear
        sleep(2)

        # save file from save as dialog as a csv in default folder
        # eg: "uia_controls.ToolbarWrapper - 'Address: C:\\Users\\zacau\\Documents\\mr107-radon-measurements', Toolbar"
        file_dir = f"{main_win.ToolBar5.wrapper_object()}"
        print(f"file dir = {file_dir}")
        if "Address: " not in file_dir:
            print(f"Address not in {file_dir}, using override")
            file_dir = "uia_controls.ToolbarWrapper - 'Address: C:\\Users\\zacau\\Documents\\mr107-radon-measurements', Toolbar"
        file_dir = file_dir.split("Address: ")[1].replace("'", "").split(",")[0]

        main_win.window(best_match='Save As Type:ComboBox').select(
            'csv (*.csv)')

        file_name = main_win.window(
            best_match='File name:ComboBox').selected_text()
        main_win.Save.click()

        file_full = os.path.join(file_dir, file_name)

        print(f"downloaded data to {file_full}")

        # wait for save dialog to finish
        sleep(2)

        # finish up and close
        main_win.Close.click()
        # app.wait_for_process_exit()

    except Exception as e:
        print(e)
        main_win.Close.click()
        raise e

    return file_dir


def clean_float(raw_float):
    # due to sensor accuracy, field values may have a '< ' in front of them
    # (eg: '< 0.8') so try to handle this
    while len(raw_float) != 0 and not raw_float[0].isdigit():
        raw_float = raw_float[1:]
    return float(raw_float)


def parse_csv(file_path):
    print(f"parsing {file_path}")
    points = []
    with open(file_path, newline='') as csvfile:
        mr_data_reader = csv.reader(csvfile, delimiter=';', quotechar='|')
        for row in mr_data_reader:
            # Series;#;Start date;Start time;Exposition;Rn activity;Temperature;Humidity;Descripton;
            # 7;1;2022.12.12;23:50;4:00;3.1;67.2;40; ;
            if len(row) < 8:
                continue
            try:
                int(row[0])
            except ValueError:
                continue

            [ser, ser_num, start_date, start_time,
                expose, rn, temp, humid, *desc] = row

            t = datetime.strptime(
                start_date + 'T' + start_time, '%Y.%m.%dT%H:%M')

            p = Point("mr107-measurements").\
                tag("room", "basement").\
                tag("series", ser).\
                field("temperature", clean_float(temp)).\
                field("humidity", clean_float(humid)).\
                field("radon", clean_float(rn)).\
                time(t)

            points.append(p)

    return points


def upload_data_to_influxdb(data_points):
    config = dotenv_values(".env")

    print(f"uploading {len(data_points)} points to influxdb {config['url']}")

    client = InfluxDBClient(
        url=config['url'],
        token=config['token'],
        org=config['org']
    )

    write_api = client.write_api(write_options=SYNCHRONOUS)

    # throws an exception upon failure
    write_api.write(bucket=config['bucket'],
                    org=config['org'], record=data_points)


def process_new_data(data_dir):
    if data_dir is None:
        return
    print(f"looking for csv files in {data_dir}")

    for entry in os.scandir(data_dir):
        if not entry.name.endswith('.csv'):
            continue

        data_points = parse_csv(entry.path)
        if data_points is None:
            continue

        upload_data_to_influxdb(data_points)

        os.rename(entry.path, entry.path + ".done")
        print(f"{entry.path}.done completed")


def main():
    data_dir = download_new_data()
    #data_dir = 'C:\\Users\\zacau\\Documents\\mr107-radon-measurements'
    process_new_data(data_dir)


if __name__ == "__main__":
    main()
