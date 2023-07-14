# Generate-to-excel
JOB TASK GENERATE DB TO EXCEL

FUNGSI :
- Membuat Console Task, untuk melakukan Generate data dari DB ke file lokal dalam bentuk file excel.
- Script auto connect dan disconnect vpn jika console job telah selesai

YANG DI BUTUHKAN :
- PHP 8+
- cisco anny connect
- SQL SERVER v18.+
- PHP Cli
- xlsxwriter.class.php

CARA PENGGUNAAN :
- Install PHP 8+ https://windows.php.net/download#php-8.1
- Install Cisco Anny Connect
- Install SQL SERVER v18.+
- Copy file user_info.txt kedalam directory C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client
- Buat Folder baru denngan nama CONSOLE_DATA pada directory D:
- Buat folder dengan nama FILE_EXCEL didalam Folder CONSOLE_DATA
- Download semua file pada github dan masukan ke dalam folder CONSOLE_DATA
- Buat Task Scheduler di PC dan arahkan Job Task ke file RUN_CONSOLE.bat atau Running langsung RUN_CONSOLE.bat
