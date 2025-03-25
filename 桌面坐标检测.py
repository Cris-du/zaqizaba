import ctypes
import time

# 定义 Windows API 获取鼠标位置
class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

GetCursorPos = ctypes.windll.user32.GetCursorPos
GetCursorPos.argtypes = [ctypes.POINTER(POINT)]
GetCursorPos.restype = ctypes.c_int

# 主循环：实时检测鼠标位置
def print_mouse_position():
    print("移动鼠标以获取实时坐标（按 Ctrl+C 停止）...")
    try:
        while True:
            # 初始化 POINT 结构体
            point = POINT()
            
            # 获取当前鼠标位置
            if GetCursorPos(ctypes.byref(point)):
                print(f"当前鼠标坐标: (x={point.x}, y={point.y})", end="\r")  # \r 用于在同一行刷新输出
            
            time.sleep(0.1)  # 每 0.1 秒刷新一次
    except KeyboardInterrupt:
        print("\n检测结束！")

# 执行检测
print_mouse_position()