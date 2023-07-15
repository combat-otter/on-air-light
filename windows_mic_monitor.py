import time
import winreg
import json
import re
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

with open("config.json") as settings_file:
    TAPO_CREDS = json.load(settings_file)


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
                        print(f"{app_name} is using the microphone")
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
                        print(f"{app_name} is using the microphone")
                        active = True
                        break

                except FileNotFoundError:
                    # reg values not found, skip
                    pass

    return active


def main():
    response = tapoPlugApi.getDeviceInfo(TAPO_CREDS)
    print(f"Using Tapo plug: {response}")

    while True:
        if is_microphone_active():
            tapoPlugApi.plugOn(TAPO_CREDS)
        else:
            tapoPlugApi.plugOff(TAPO_CREDS)

        time.sleep(2)


if __name__ == "__main__":
    main()
