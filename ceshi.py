playwright codegen https://www.e-custody.com/#/


pyinstaller --clean dabao.spec
pyinstaller --clean --noconfirm dabao.spec

#清除build
rmdir /S /Q build

#浏览器内核一起打包
pyinstaller --noconsole --onefile --add-data "config.txt;." --add-data "playwright-browsers;playwright-browsers" main.py


if hasattr(sys, "_MEIPASS"):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.abspath(os.path.dirname(__file__))



#config和内核手动复制 只打包py：好像可以清理build后打包
pyinstaller --clean --noconsole --onefile --add-data "config.txt;." bankdownloader.py



if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)  # .exe 所在目录
else:
    base_dir = os.path.abspath(os.path.dirname(__file__))