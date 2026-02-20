"""Windows Test MCP 使用示例.

此脚本演示如何使用各个工具进行Windows桌面自动化测试。
"""

import time


def demo_screenshot():
    """演示截图功能."""
    print("=== 截图演示 ===")

    # 注意: 这里的工具调用需要在MCP环境中执行
    # 以下只是伪代码示例

    # 截取整个屏幕
    # result = screenshot_capture()
    # if result.success:
    #     print(f"截图成功，base64长度: {len(result.base64_image)}")

    # 截取区域 (左上角100,100，大小800x600)
    # result = screenshot_region(x=100, y=100, width=800, height=600)
    # if result.success:
    #     print(f"区域截图成功")

    print("截图演示完成\n")


def demo_keyboard():
    """演示键盘操作."""
    print("=== 键盘操作演示 ===")

    # 等待2秒，让用户切换到目标窗口
    print("2秒后开始键盘演示...")
    time.sleep(2)

    # 输入文本
    # keyboard_type(text="Hello, Windows Test MCP!", interval=0.01)

    # 按Enter键
    # keyboard_press(key="enter")

    # 按Tab键
    # keyboard_press(key="tab")

    # 组合键示例: Ctrl+A (全选)
    # keyboard_down(key="ctrl")
    # keyboard_press(key="a")
    # keyboard_up(key="ctrl")

    # 组合键示例: Ctrl+C (复制)
    # keyboard_down(key="ctrl")
    # keyboard_press(key="c")
    # keyboard_up(key="ctrl")

    print("键盘操作演示完成\n")


def demo_mouse():
    """演示鼠标操作."""
    print("=== 鼠标操作演示 ===")

    # 等待2秒
    print("2秒后开始鼠标演示...")
    time.sleep(2)

    # 获取当前鼠标位置
    # result = mouse_get_position()
    # print(f"当前鼠标位置: {result.current_position}")

    # 移动鼠标到屏幕中心
    # screen = get_screen_size()
    # center_x = screen.width // 2
    # center_y = screen.height // 2
    # mouse_move(x=center_x, y=center_y, duration=0.5)

    # 左键点击
    # mouse_click(button="left")

    # 右键点击
    # mouse_click(button="right")

    # 双击
    # mouse_click(clicks=2)

    # 滚动滚轮 (正数向上，负数向下)
    # mouse_scroll(clicks=5)

    print("鼠标操作演示完成\n")


def demo_workflow():
    """演示一个完整的工作流程."""
    print("=== 完整工作流程演示 ===")
    print("场景: 打开记事本，输入文本，保存")

    # 步骤1: 截图查看当前状态
    # screenshot_capture(filename="step1_initial")

    # 步骤2: 移动鼠标并点击开始菜单
    # mouse_move(x=50, y=1050, duration=0.5)
    # mouse_click()

    # 步骤3: 输入"notepad"
    # keyboard_type(text="notepad", interval=0.01)
    # keyboard_press(key="enter")

    # 步骤4: 等待记事本打开
    # time.sleep(1)

    # 步骤5: 输入内容
    # keyboard_type(text="这是一个自动化测试示例。", interval=0.01)

    # 步骤6: 保存 (Ctrl+S)
    # keyboard_down(key="ctrl")
    # keyboard_press(key="s")
    # keyboard_up(key="ctrl")

    # 步骤7: 输入文件名
    # keyboard_type(text="test_document", interval=0.01)
    # keyboard_press(key="enter")

    print("工作流程演示完成\n")


if __name__ == "__main__":
    print("Windows Test MCP 使用示例")
    print("=" * 40)
    print()

    # 演示各个功能
    demo_screenshot()
    demo_keyboard()
    demo_mouse()
    demo_workflow()

    print("=" * 40)
    print("所有演示完成!")
    print()
    print("注意: 这些工具需要在MCP环境中调用")
    print("请查看 README.md 了解如何在Claude Desktop中配置使用")
