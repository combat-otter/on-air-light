import time
import winreg
import json
import re
import logging
import contextlib
from PyP100 import PyP100

"""
Monitors Windows apps for microphone usage, switches plug on if mic is in use.

Requires a local config.json file with the following content:

{
    "tapoIp": "DEVICE IP",
    "tapoEmail": "ENTER YOUR TP-LINK EMAIL",
    "tapoPassword": "ENTER YOUR TP-LINK PWD"
}
"""

# console logging setup
logging.basicConfig(format="%(asctime)s %(levelname)-8s %(message)s", level=logging.INFO, datefmt="%Y-%m-%d %H:%M:%S")

# config file currently only used for Tapo plug settings
with open("config.json") as config_file:
    TAPO_CREDS = json.load(config_file)


def get_app_using_mic():
    app_name = None

    with winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone",
    ) as root_key:
        (sub_key_count, value_count, last_modified) = winreg.QueryInfoKey(root_key)
        for i in range(sub_key_count):
            sub_key_name = winreg.EnumKey(root_key, i)

            with winreg.OpenKey(root_key, sub_key_name) as sub_key:
                try:
                    (last_used_time_stop, _) = winreg.QueryValueEx(sub_key, "LastUsedTimeStop")

                    if last_used_time_stop == 0:
                        app_name = re.sub(r"_([a-z0-9])+$", "", sub_key_name)
                        break

                except FileNotFoundError:
                    # reg values not found, skip
                    pass

    if app_name:
        return app_name

    with winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone\NonPackaged",
    ) as root_key:
        (sub_key_count, value_count, last_modified) = winreg.QueryInfoKey(root_key)

        for i in range(sub_key_count):
            sub_key_name = winreg.EnumKey(root_key, i)

            with winreg.OpenKey(root_key, sub_key_name) as sub_key:
                try:
                    (last_used_time_stop, _) = winreg.QueryValueEx(sub_key, "LastUsedTimeStop")

                    if last_used_time_stop == 0:
                        app_name = re.sub(r"([\w:# \(\)\-])+#", "", sub_key_name)
                        break

                except FileNotFoundError:
                    # reg values not found, skip
                    pass

    return app_name


# add context manager to the plug api
@contextlib.contextmanager
def get_plug_instance():
    plug_context = PyP100.P100(
        address=TAPO_CREDS["tapoIp"],
        email=TAPO_CREDS["tapoEmail"],
        password=TAPO_CREDS["tapoPassword"],
    )
    yield plug_context
    del plug_context


def main():
    active_seconds = 0
    inactive_seconds = 0
    hysteresis = 10

    switched_on = False

    while True:
        try:
            time.sleep(1)

            tick = int(time.time()) % 30
            if tick == 0:
                with get_plug_instance() as plug:
                    response = plug.getDeviceInfo()
                    logging.debug(response)
                    logging.info(f"Plug powered: {response['device_on']}")
                    switched_on = response["device_on"]

            app_using_mic = get_app_using_mic()
            if app_using_mic:
                if active_seconds <= hysteresis:
                    logging.info(f"{app_using_mic} is using the microphone")

                active_seconds += 1
                inactive_seconds = 0
            else:
                active_seconds = 0
                inactive_seconds += 1

            if active_seconds > hysteresis and not switched_on:
                logging.info("Switching plug on")
                with get_plug_instance() as plug:
                    plug.turnOn()
                    switched_on = plug.getDeviceInfo()["device_on"]

            elif inactive_seconds > hysteresis and switched_on:
                logging.info("Switching plug off")
                with get_plug_instance() as plug:
                    plug.turnOff()
                    switched_on = plug.getDeviceInfo()["device_on"]

        except KeyboardInterrupt:
            break

        except Exception as e:
            logging.error(e)
            continue


if __name__ == "__main__":
    main()
