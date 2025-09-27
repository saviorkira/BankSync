from playwright.sync_api import sync_playwright

# 启动 Playwright
with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)  # headless=False 方便观察
    context = browser.new_context()
    page = context.new_page()

    # 存储捕获的网络请求
    requests = []

    # 监听网络请求
    def log_request(request):
        if 'basic_info/select' in request.url:  # 过滤目标 API
            requests.append({
                'url': request.url,
                'method': request.method,
                'headers': request.headers,
                'post_data': request.post_data
            })
        print(f"Captured request: {request.url}")

    page.on("request", log_request)

    # 访问页面
    page.goto("http://172.16.10.185:8443/ks/")

    # 定位按钮并点击（假设你已找到按钮的 selector）
    # 替换为实际的按钮选择器，例如：
    # - CSS 选择器：如 '#seekBtn'（从 .xwl 文件看到查询按钮的 itemId 是 seekBtn）
    # - 文本匹配：如 page.get_by_text("查询")
    button_selector = '#seekBtn'  # 请根据 F12 或 Playwright Inspector 确认
    page.locator(button_selector).click()

    # 等待网络请求完成
    page.wait_for_timeout(5000)  # 等待 5 秒，确保请求完成

    # 打印捕获的请求
    for req in requests:
        print("Captured API Request:")
        print(f"URL: {req['url']}")
        print(f"Method: {req['method']}")
        print(f"Headers: {req['headers']}")
        print(f"Post Data: {req['post_data']}")

    # 关闭浏览器
    browser.close()