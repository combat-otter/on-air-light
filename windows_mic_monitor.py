import time
import winreg
import json
import re
import logging
from tapo_plug import tapoPlugApi

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


def is_microphone_active():
    active = False

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
                        logging.info(f"{app_name} is using the microphone")
                        active = True
                        break

                except FileNotFoundError:
                    # reg values not found, skip
                    pass

    if active:
        return active

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
                        logging.info(f"{app_name} is using the microphone")
                        active = True
                        break

                except FileNotFoundError:
                    # reg values not found, skip
                    pass

    return active


def main():
    response = tapoPlugApi.getDeviceInfo(TAPO_CREDS)
    print(f"Using Tapo plug: {response}")
    plug_on = False

    while True:
        if is_microphone_active():
            if not plug_on:
                logging.info("Switching plug on")
                response = json.loads(tapoPlugApi.plugOn(TAPO_CREDS))
                plug_on = response["error_code"] == 0

        else:
            if plug_on:
                logging.info("Switching plug off")
                response = json.loads(tapoPlugApi.plugOff(TAPO_CREDS))
                plug_on = not response["error_code"] == 0

        time.sleep(5)


if __name__ == "__main__":
    main()
