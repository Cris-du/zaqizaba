import pyperclip
import keyboard
import time
import ctypes

# ================== 参数配置 ==================
# 定义鼠标点击的坐标
STEP_1_POSITION = (1253, 723)  # 第一次点击的位置
STEP_2_POSITION = (262, 440)    # 第二次点击的位置
STEP_3_POSITION = (564, 464)    # 第三次点击的位置
STEP_4_POSITION = (435, 684)  # 第四次点击的位置
STEP_5_POSITION = (1299, 727)  # 第五次点击的位置

# 新增两个监控光标的坐标
READ_CURSOR_POSITION_1 = (1298, 726)  # 监控坐标 1
READ_CURSOR_POSITION_2 = (754, 567)  # 监控坐标 2

# 新增检测光标类型的坐标
NEW_POSITION = (413, 519)  # 新增的检测坐标

# 定义要输入的固定文本
INPUT_2_TEXT = "详细阅读这个pdf，从摘要到介绍到研究方法到研究结果到研究结论，让我弄懂这篇文章在做什么。要求以表格输出，格式为第一行是表头，每列是每个标题的内容，分别有研究环境	深度	地理位置	状态	研究对象	年份	领域	结论	主要工作	这里要注意的是：1.研究环境分为海洋热液、海洋冷泉、海洋深渊，如果不是以上三者，则看是不是海洋环境，如果不是海洋环境，则根据文章内容来定。2.深度分为深海或浅海，如果两者都有，则输出为深海；浅海。3.地理位置按照方法中的描述输出，只需要输出地理位置名称，不要输出编号这种个性化的内容，也不要宽泛，要具体，如果有多个地理位置，用；符号分隔。4.状态是研究样品的状态，按照方法的描述输出，如果有多个状态，用；符号分隔。5.研究对象分为病毒、微生物、动物、植物等，这个按照文章描述来确定，同样如果有多个，用；符号分隔。6.文章年份按照文章信息来确定。7.文章的领域需要稍微细致一些，不要太宽泛，基于文章做出的真实结果和所做的工作来判断，不要给出病毒学、极端病毒学这种很宽泛的信息，我需要的是更具体一些的，而不是根据文章推测的信息或者假设等信息判断。8.文章的主要结果和结论，原则是清晰且完整，清晰指的是不要有任何推测性的描述，比如可能、也许、展望这种字眼，文章中推测性描述也都忽略掉，只管他做了哪些具体的工作得到的结果和结论，完整指的是要详细描述它的结果和结论，注意是结果和结论，既要概况也要分别描述，尽量一句话概况。9.文章的主要工作做了什么，也就是文章真实做的实验，不同的实验用；列出即可。表格除了第一行表头外，只需要一行来输出结果"

# 鼠标事件常量
MOUSEEVENTF_LEFTDOWN = 0x0002  # 左键按下
MOUSEEVENTF_LEFTUP = 0x0004    # 左键释放

# 光标类型常量
IDC_ARROW = 32512       # 默认箭头光标
IDC_HAND = 32649        # 手型光标
IDC_NO = 32648          # 禁止图标光标

# 定义 .txt 文件的绝对路径
INPUT_FILE_PATH = r"D:\桌面\test.txt"  # 输入文件路径

# ================== Windows API 获取光标信息 ==================
class CURSORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_uint32),
        ("flags", ctypes.c_uint32),
        ("hCursor", ctypes.c_void_p),
        ("ptScreenPos", ctypes.c_long * 2)
    ]

# 加载必要的 Windows API 函数
user32 = ctypes.windll.user32
GetCursorInfo = user32.GetCursorInfo
GetCursorInfo.argtypes = [ctypes.POINTER(CURSORINFO)]
GetCursorInfo.restype = ctypes.c_int

LoadCursorW = user32.LoadCursorW
LoadCursorW.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
LoadCursorW.restype = ctypes.c_void_p

# 加载常见光标的句柄
arrow_cursor = LoadCursorW(0, IDC_ARROW)
hand_cursor = LoadCursorW(0, IDC_HAND)
no_cursor = LoadCursorW(0, IDC_NO)

# ================== 函数定义 ==================
def move_and_click(x, y):
    ctypes.windll.user32.SetCursorPos(x, y)  # 移动鼠标到指定位置
    time.sleep(0.1)  # 等待鼠标移动完成
    ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)  # 按下左键
    time.sleep(0.1)
    ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)    # 释放左键
    time.sleep(0.1)

def type_text(text):
    pyperclip.copy(text)
    time.sleep(0.1)  # 等待剪贴板更新完成
    keyboard.press_and_release('ctrl+v')
    time.sleep(0.1)  # 等待粘贴完成

def monitor_cursor(position1, position2):
    print(f"开始监控光标类型，每隔 3 秒切换一次坐标...")
    while True:
        for idx, position in enumerate([position1, position2], start=1):
            x, y = position
            ctypes.windll.user32.SetCursorPos(x, y)
            time.sleep(3)  # 等待鼠标移动完成
            cursor_info = CURSORINFO()
            cursor_info.cbSize = ctypes.sizeof(CURSORINFO)
            if GetCursorInfo(ctypes.byref(cursor_info)):
                current_cursor = cursor_info.hCursor
                if current_cursor == arrow_cursor:
                    cursor_type = "默认箭头光标 (IDC_ARROW)"
                elif current_cursor == hand_cursor:
                    cursor_type = "手型光标 (IDC_HAND)"
                elif current_cursor == no_cursor:
                    cursor_type = "禁止图标光标 (IDC_NO)"
                    print(f"在坐标 {position} 检测到光标类型: {cursor_type}")
                    print("检测到禁止图标光标，停止监控。")
                    return True
                else:
                    cursor_type = f"未知光标类型: {current_cursor}"
                print(f"当前位置 {position} 的光标类型: {cursor_type}")
        print("未检测到禁止图标光标，继续轮询...")

# 新增函数：检测 NEW_POSITION 的光标类型并处理
def check_hand_cursor_and_loop(new_position, step_5_position):
    """
    检测 NEW_POSITION 的光标类型是否为手型，并根据结果执行循环操作。
    """
    print(f"开始检测 {new_position} 的光标类型...")
    while True:
        # 移动鼠标到 NEW_POSITION
        ctypes.windll.user32.SetCursorPos(*new_position)
        time.sleep(0.5)  # 等待鼠标移动完成
        
        # 获取当前光标信息
        cursor_info = CURSORINFO()
        cursor_info.cbSize = ctypes.sizeof(CURSORINFO)
        if GetCursorInfo(ctypes.byref(cursor_info)):
            current_cursor = cursor_info.hCursor
            
            # 判断光标类型
            if current_cursor == hand_cursor:
                print(f"检测到手型光标，在 {new_position} 处。")
                # 回到 STEP_5_POSITION 并点击
                print(f"回到 {step_5_position} 并点击...")
                move_and_click(*step_5_position)
                time.sleep(1)  # 等待点击完成
            else:
                print(f"光标类型不再是手型，退出循环。")
                break

# ================== 主函数 ==================
def automate_task():
    try:
        # 从文件中读取 INPUT_1_TEXT 列表
        with open(INPUT_FILE_PATH, "r", encoding="utf-8") as file:
            input_texts = [line.strip() for line in file if line.strip()]
        
        print(f"共读取到 {len(input_texts)} 条输入内容，开始自动化任务...")
        
        # 遍历每一条输入内容
        for idx, input_text in enumerate(input_texts, start=1):
            print(f"\n开始处理第 {idx} 条输入内容: {input_text}")
            
            # 步骤 1：将鼠标移动到 STEP_1_POSITION 并点击
            print(f"步骤 1：移动鼠标到 {STEP_1_POSITION} 并点击...")
            move_and_click(*STEP_1_POSITION)
            time.sleep(0.1)
            
            # 步骤 2：将鼠标移动到 STEP_2_POSITION 并点击
            print(f"步骤 2：移动鼠标到 {STEP_2_POSITION} 并点击...")
            move_and_click(*STEP_2_POSITION)
            time.sleep(0.1)
            
            # 步骤 3：输入文本 INPUT_1_TEXT
            print(f'步骤 3：输入文本 "{input_text}"...')
            type_text(input_text)
            time.sleep(0.1)
            
            # 步骤 4：将鼠标移动到 STEP_3_POSITION 并点击
            print(f"步骤 4：移动鼠标到 {STEP_3_POSITION} 并点击...")
            move_and_click(*STEP_3_POSITION)
            time.sleep(0.1)
            
            # 步骤 5：将鼠标移动到 STEP_4_POSITION 并点击
            print(f"步骤 5：移动鼠标到 {STEP_4_POSITION} 并点击...")
            move_and_click(*STEP_4_POSITION)
            time.sleep(0.1)
            
            # 步骤 6：输入文本 INPUT_2_TEXT
            print(f'步骤 6：输入文本 "{INPUT_2_TEXT}"...')
            type_text(INPUT_2_TEXT)
            time.sleep(3)
            
            # 步骤 7：将鼠标移动到 STEP_5_POSITION 并点击
            print(f"步骤 7：移动鼠标到 {STEP_5_POSITION} 并点击...")
            move_and_click(*STEP_5_POSITION)
            time.sleep(0.1)
            
            # 新增步骤：检测 NEW_POSITION 的光标类型
            print(f"新增步骤：检测 {NEW_POSITION} 的光标类型...")
            check_hand_cursor_and_loop(NEW_POSITION, STEP_5_POSITION)
            
            # 步骤 8：监控两个坐标点的光标类型
            print(f"步骤 8：监控两个坐标点 {READ_CURSOR_POSITION_1} 和 {READ_CURSOR_POSITION_2} 的光标类型...")
            monitor_cursor(READ_CURSOR_POSITION_1, READ_CURSOR_POSITION_2)
        
        print("\n所有输入内容处理完毕，自动化任务完成！")
    
    except Exception as e:
        print(f"发生错误: {e}")

# 执行自动化任务
automate_task()