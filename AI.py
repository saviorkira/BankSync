import os
from openai import OpenAI
import threading
from utils import log

# 初始化 OpenAI 客户端
try:
    client = OpenAI(
        api_key="sk-260cfb7ab40440e695dbe9eda9aa9a4d",
        base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
    )
except Exception as e:
    client = None
    log(f"OpenAI 客户端初始化失败: {str(e)}", os.path.abspath(os.path.dirname(__file__)))

# AI 消息列表
ai_messages = []

def update_ai_output(ai_output, selected_index, project_root):
    try:
        ai_output.value = "".join(ai_messages)
        if selected_index.current == 2 and hasattr(ai_output, 'page') and ai_output.page is not None:
            ai_output.update()
            ai_output.page.scroll_to(key="ai_output", duration=500)
            log("AI: ai_output 更新成功: AI页面已渲染", project_root)
        else:
            log("AI: ai_output 未更新: 未在AI页面或未添加到页面", project_root)
    except Exception as e:
        log(f"AI: ai_output 更新失败: {str(e)}", project_root)

def send_ai_message(e, ai_input, ai_output, update_log, selected_index, project_root):
    try:
        update_log("AI: 收到发送消息请求")
        if not ai_input.value.strip():
            update_log("AI: 错误: 输入消息不能为空")
            return
        user_message = ai_input.value.strip()
        ai_messages.append(f"用户: {user_message}\n")
        update_log(f"AI: 已记录用户消息: {user_message}")
        update_ai_output(ai_output, selected_index, project_root)

        ai_input.value = ""
        if hasattr(ai_input, 'page') and ai_input.page is not None:
            ai_input.update()
            update_log("AI: ai_input 更新: 输入框已清空")
        else:
            update_log("AI: ai_input 未更新: 未添加到页面")

        def ai_worker():
            try:
                if not client:
                    ai_messages.append("AI: 错误: OpenAI 客户端未初始化\n")
                    update_ai_output(ai_output, selected_index, project_root)
                    update_log("AI: 错误: OpenAI 客户端未初始化")
                    return
                if "supergrok" in user_message.lower():
                    ai_messages.append("AI: 关于 SuperGrok 的定价和使用限制，请访问 https://x.ai/grok 获取详细信息。\n")
                    update_log("AI: 检测到 SuperGrok 查询，返回定价指引")
                else:
                    completion = client.chat.completions.create(
                        model="qwen-plus",
                        messages=[
                            {"role": "system", "content": "You are a helpful assistant."},
                            {"role": "user", "content": user_message},
                        ],
                    )
                    response = completion.choices[0].message.content
                    ai_messages.append(f"AI: {response}\n")
                    update_log("AI: 成功接收 AI 回复")
                update_ai_output(ai_output, selected_index, project_root)
                update_log("AI: 回复处理完成")
            except Exception as ex:
                ai_messages.append(f"AI: 错误: {str(ex)}\n")
                update_ai_output(ai_output, selected_index, project_root)
                update_log(f"AI: 回复失败: {str(ex)}")

        threading.Thread(target=ai_worker, daemon=True).start()
        update_log("AI: 已启动 AI 处理线程")
    except Exception as e:
        update_log(f"AI: 发送消息失败: {str(e)}")