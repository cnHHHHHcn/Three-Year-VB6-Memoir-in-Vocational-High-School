netsh interface ip set address name="本地连接" source=dhcp
netsh interface set interface "本地连接" admin=disable
netsh interface set interface "本地连接" admin=enabled