netsh interface ip set address name="本地连接" static 192.168.1.16 255.255.255.0 
netsh interface set interface "本地连接" admin=disable
netsh interface set interface "本地连接" admin=enabled