import webview
class Api:
    def pick_file(self, extensions):
        return ["test.xlsx"]
    def pick_directory(self):
        return "D:\\Test"
api = Api()
window = webview.create_window('测试窗口', 'https://www.example.com', width=600, height=700, resizable=False, minimizable=True, maximizable=False)
webview.start(globals={'pywebview': {'api': api}})
print("窗口已启动")