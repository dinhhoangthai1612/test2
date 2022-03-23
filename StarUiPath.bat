netsh interface set interface Ethernet disable

timeout 5
netsh interface set interface Ethernet enable

timeout 10
::EAP SYSTEM
net use \\10.246.194.71 /user:VNMSALESSUP\Admin Password@123
net use \\10.247.194.66 /user:YKKHaNam\VIS Password@123

::ROBOT NHON TRACH
net use \\10.246.192.241 /user:Ykkvietnam\robotics Password@123
net use \\10.246.194.28 /user:Ykkvietnam\robotics Password@123
net use \\10.246.194.33 /user:Ykkvietnam\robotics Password@123
net use \\10.246.194.36 /user:Ykkvietnam\robotics Password@123

::DEV NT
net use \\10.246.194.73 /user:Ykkvietnam\robotics Password@123
net use \\10.246.194.74 /user:Ykkvietnam\robotics Password@123
net use \\10.246.194.75 /user:Ykkvietnam\robotics Password@123

::ROBOT HCM
net use \\10.246.194.27 /user:Ykkvietnam\robotics Password@123
net use \\10.246.192.236 /user:Ykkvietnam\robotics Password@123
net use \\10.246.192.216 /user:Ykkvietnam\robotics Password@123
net use \\10.246.192.235 /user:Ykkvietnam\robotics Password@123

::DEV HCM
net use \\10.246.192.231 /user:Ykkvietnam\robotics Password@123
net use \\10.246.192.232 /user:Ykkvietnam\robotics Password@123
net use \\10.246.192.233 /user:Ykkvietnam\robotics Password@123
net use \\10.246.192.249 /user:Ykkvietnam\robotics Password@123
net use \\10.246.192.248 /user:Ykkvietnam\robotics Password@123

::COMPLEO
net use \\10.246.194.31 /user:Ykkvietnam\Administrator YKK@dmin123
net use \\10.247.194.1 /user:ROBOT PASSWORD
net use \\10.246.194.1 /user:ROBOTNT ROBOTNT123

::Waiting for star Server Manager Finish
taskkill /IM "UiPath.Executor.exe" /F
timeout 35

::"C:\Program Files (x86)\UiPath\Studio\UiRobot.exe" -file "D:\Server\YkkNT\IBM I SIGNON.xaml"

::ipconfig | findstr /C:"10.246.194.75" >nul && (
::    "C:\Program Files (x86)\UiPath\Studio\UiRobot.exe" -file "D:\Server\YkkNT\MailServer.xaml"
::) || (
::    "C:\Program Files (x86)\UiPath\Studio\UiRobot.exe" -file "D:\Server\YkkNT\MainServer_Automatic.xaml"
::)

start H:
start Z:
::start \\10.247.194.230\ykkhnf_data$
"C:\Program Files (x86)\UiPath\Studio\UiRobot.exe" -file "D:\Server\YkkNT\MainServer_Automatic.xaml"

pause