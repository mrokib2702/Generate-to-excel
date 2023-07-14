@echo off
C:
cd "C:/Program Files (x86)/Cisco/Cisco AnyConnect Secure Mobility Client"

vpncli.exe disconnect
vpncli.exe -s < user_info.txt

exit


D:
cd "D:/CONSOLE_DATA"
PAUSE

