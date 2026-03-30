__version__ = "2.17.1"
__author__ = "随机点名程序"

# 随机点名程序 v2.17
# 导入必要的库
from random import randint            # 随机数生成器
from tkinter import filedialog        # 文件选择对话框
import subprocess                     # ┬系统命令执行
import os                             # ┤
import sys                            # ┘
import tkinter as t                   # GUI界面库
import ctypes                         # 调用Windows API
from ctypes import wintypes           # Windows数据类型


# 获取屏幕尺寸
屏幕 = t.Tk()
屏幕宽度 = 屏幕.winfo_screenwidth()
屏幕高度 = 屏幕.winfo_screenheight()
屏幕.destroy()  # 销毁临时窗口

# 常量定义
# 计算窗口初始位置（窗口中心点放在屏幕下方1/3左侧）
展开窗口宽度 = 600
展开窗口高度 = 200
收起窗口宽度 = 30
收起窗口高度 = 60
初始x = 展开窗口宽度 / 2  # 窗口中心点x坐标（左侧对齐）
初始y = 屏幕高度 * 2 / 3  # 窗口中心点y坐标（屏幕下方1/3处）
# 转换为窗口顶点坐标
窗口顶点x = 初始x - 展开窗口宽度 / 2
窗口顶点y = 初始y - 展开窗口高度 / 2

# 全局变量初始化
姓名列表, 概率列表, 临时姓名列表, 临时概率列表 = [], [], [], []
当前名字 = "点击下方按钮开始"
收缩帧率 = 20  # 收起动画所需要的帧数
启动_关闭帧数 = 80000  # 启动/关闭窗口动画所需要的帧数
透明度 = 0.8  # 窗口透明度，默认80%

# 自定义异常类
class OpenError(Exception):
    pass

# 使用Windows互斥体实现单实例运行（不生成文件）
def 创建互斥体检查单实例():
    """使用Windows Mutex确保只有一个程序实例运行"""
    # 定义Windows API函数
    CreateMutex = ctypes.windll.kernel32.CreateMutexW
    CreateMutex.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
    CreateMutex.restype = wintypes.HANDLE
    
    GetLastError = ctypes.windll.kernel32.GetLastError
    
    ERROR_ALREADY_EXISTS = 183
    
    # 创建一个命名互斥体（使用唯一标识符）
    mutex_name = "RandomRollCallPicker_v2.17_SingleInstance"
    mutex = CreateMutex(None, False, mutex_name)
    
    if mutex == 0:
        raise OpenError("无法创建互斥体，程序启动失败。")
    
    if GetLastError() == ERROR_ALREADY_EXISTS:
        # 互斥体已存在，说明程序已在运行
        ctypes.windll.kernel32.CloseHandle(mutex)
        raise OpenError("已有程序正在运行。\nIt's running now.")
    
    return mutex

# 保存互斥体句柄，防止被垃圾回收
互斥体句柄 = 创建互斥体检查单实例()

def 加载默认数据文件():
    """加载默认数据文件"""
    global 当前名字
    默认文件路径 = "姓名及概率.txt"
    
    if not os.path.exists(默认文件路径):
        当前名字 = "未找到文件"
        return False
    
    try:
        with open(默认文件路径, encoding='utf-8') as 文件:
            内容 = 文件.read()
        
        数据行 = 内容.strip().split("\n")
        
        # 验证文件格式
        if not 数据行 or 数据行[0] != "姓名,概率":
            raise ValueError("数据文件格式错误：第一行应为'姓名,概率'")
        
        # 解析数据
        for 行 in 数据行[1:]:
            if 行.strip():  # 跳过空行
                部分 = 行.split(",")
                if len(部分) >= 2:
                    姓名列表.append(部分[0].strip())
                    概率列表.append(int(float(部分[1])))
        
        return True
    except Exception as e:
        当前名字 = f"文件读取错误: {str(e)}"
        return False

# 加载默认数据文件
加载默认数据文件()
    
def 读取数据文件(文件路径):
    """读取指定路径的数据文件并更新全局变量"""
    global 姓名列表, 概率列表
    
    if not os.path.exists(文件路径):
        return False
    
    try:
        with open(文件路径, encoding='utf-8') as 文件:
            数据内容 = 文件.read()
        
        数据行 = 数据内容.strip().split("\n")
        
        # 验证文件格式
        if not 数据行 or 数据行[0] != "姓名,概率":
            raise ValueError("数据文件格式错误：第一行应为'姓名,概率'")
        
        # 清空原有数据
        姓名列表.clear()
        概率列表.clear()
        
        # 解析新数据
        for 行 in 数据行[1:]:
            if 行.strip():  # 跳过空行
                部分 = 行.split(",")
                if len(部分) >= 2:
                    姓名列表.append(部分[0].strip())
                    概率列表.append(int(float(部分[1])))
        
        # 重新初始化概率数据
        初始化()
        return True
    except Exception:
        return False

def 选择文件():
    """打开文件选择对话框并读取选中的数据文件"""
    global 当前文件路径
    文件路径 = filedialog.askopenfilename(
        title="选择数据文件",
        filetypes=[("文本文件", "*.txt")],
        initialdir="."  # 当前目录
    )
    
    if 文件路径:
        if 读取数据文件(文件路径):
            # 更新界面显示
            初始化所有组件()
            # 显示成功消息
            名字标签.config(text="数据文件加载成功！", font=名字字体)
            # 更新文件路径文本
            当前文件路径 = 文件路径
            文件路径文本.config(text=当前文件路径)
        else:
            名字标签.config(text="文件读取失败，请检查格式", font=名字字体)
            # 更新文件路径文本
            当前文件路径 = "无"
            文件路径文本.config(text=当前文件路径)

def 开启动画():
    """设置透明度(递增)"""
    for 帧数 in range(启动_关闭帧数):
        窗口.wm_attributes("-alpha", (透明度 / 启动_关闭帧数) * 帧数)
        窗口.update()

def 关闭动画():
    """设置透明度(递减)"""
    for 帧数 in range(启动_关闭帧数):
        窗口.wm_attributes("-alpha", 透明度 - (透明度 / 启动_关闭帧数) * 帧数)
        窗口.update()

def 关闭():
    """关闭应用程序并释放互斥体"""
    global 互斥体句柄
    if 互斥体句柄:
        ctypes.windll.kernel32.CloseHandle(互斥体句柄)
        互斥体句柄 = None
    # 执行动画
    关闭动画()
    窗口.quit()

def 重启():
    """重启应用程序"""
    关闭()
    # 重启应用程序
    if getattr(sys, 'frozen', False):
        # 从可执行文件运行
        subprocess.Popen([sys.executable])
    else:
        # 从脚本运行
        script_path = os.path.abspath(__file__)
        subprocess.Popen([sys.executable, script_path])

def 移动(point):
    """拖动窗口功能"""
    global 隐藏状态, y坐标, 鼠标按下位置
    
    if 隐藏状态:
        # 只计算y方向的偏移量（忽略x方向）
        偏移y = point.y_root - 鼠标按下位置[1]
        
        # 获取当前窗口位置
        当前x = 窗口.winfo_x()
        当前y = 窗口.winfo_y()
        
        # 计算新的y坐标，并限制移动范围
        新y = 当前y + 偏移y
        
        # 限制移动范围：程序中心点不可移出屏幕
        # 窗口中心y坐标 = 新y + 展开窗口高度/2
        窗口中心y = 新y + 展开窗口高度 / 2
        最小y = 展开窗口高度 / 2  # 中心点最低位置（顶部边界）
        最大y = 屏幕高度  # 中心点最高位置（底部边界）
        
        # 应用范围限制
        if 窗口中心y < 最小y:
            新y = 最小y - 展开窗口高度 / 2
        elif 窗口中心y > 最大y:
            新y = 最大y - 展开窗口高度 / 2
        
        # 设置新的窗口位置（只改变y坐标，x坐标保持不变）
        窗口.geometry(f"+{当前x}+{int(新y)}")
        
        # 更新鼠标按下位置为当前位置
        鼠标按下位置 = (point.x_root, point.y_root)

def 鼠标按下(event):
    """记录鼠标按下时的位置"""
    global 鼠标按下位置
    鼠标按下位置 = (event.x_root, event.y_root)

def 复制列表(传入列表):
    """深拷贝列表"""
    return 传入列表.copy()

def 初始化():
    """初始化概率数据，同比化为整数"""
    import decimal
    
    # 找出概率中的最大小数位数
    最大小数位数 = 0
    for 值 in 概率列表:
        小数位数 = abs(decimal.Decimal(str(值)).as_tuple().exponent)
        最大小数位数 = max(最大小数位数, 小数位数)
    
    # 将概率转换为整数（避免浮点数精度问题）
    if 最大小数位数 > 0:
        倍数 = 10 ** 最大小数位数
        for 索引 in range(len(概率列表)):
            概率列表[索引] = int(概率列表[索引] * 倍数)

def 显示所有初始组件():
    """显示主界面所有组件"""
    切换模式按钮.place(relx=0, rely=1, anchor="sw")
    更多.place(relx=1, rely=1, anchor="se")
    隐藏按钮.place(relx=1, rely=0.5, anchor="e")
    标签.pack()
    名字标签.pack()
    更新按钮.pack()
    查看概率按钮.place(relx=0.5, rely=1, anchor="s")

def 隐藏所有组件():
    """隐藏所有界面组件"""
    关闭按钮.place_forget()
    名字标签.pack_forget()
    查看概率按钮.place_forget()
    隐藏按钮.place_forget()
    标签.pack_forget()
    更新按钮.pack_forget()
    更多.place_forget()
    切换模式按钮.place_forget()
    重启按钮.place_forget()
    选择文件按钮.place_forget()
    文件路径文本.place_forget()
    字体大小文本.place_forget()
    字体大小滑块.place_forget()
    窗口化按钮.place_forget()

def 初始化所有组件():
    """初始化界面组件状态"""
    global 隐藏状态, 正在查看概率
    隐藏状态 = False
    正在查看概率 = False
    隐藏所有组件()
    显示所有初始组件()
    # 配置按钮文本和功能
    关闭按钮.config(text="x", command=窗口.quit)
    切换模式按钮.config(text=f"当前模式:{当前模式}", command=切换模式)
    重启按钮.config(text="重启", command=重启)
    更多.config(text="更多...", command=更多选项显示)
    隐藏按钮.config(text="<", command=隐藏)
    名字标签.config(text=当前名字, font=名字字体)
    更新按钮.config(text="按钮", command=更新)
    查看概率按钮.config(text="查看概率", command=查看概率)
    选择文件按钮.config(text="选择数据文件", command=选择文件)

def 随机函数(传入姓名列表, 传入概率列表):
    """根据概率随机选择姓名"""
    概率总数 = sum(传入概率列表)
    当前概率数 = randint(1, 概率总数)
    累计概率 = 0
    
    # 根据概率区间选择姓名
    for 索引, 当前项 in enumerate(传入概率列表):
        累计概率 += 当前项
        if 当前概率数 <= 累计概率:
            return (传入姓名列表[索引], 索引)  # 返回姓名和索引

def 随机():
    """根据当前模式进行随机选择"""
    global 当前名字, 临时姓名列表, 临时概率列表
    
    if 当前模式 == "每一轮不重复(推荐)":
        # 模式1：一轮内不重复，直到所有概率>0的人都被点到
        概率低人数 = sum(1 for i in 概率列表 if i <= 0)
        # 如果临时列表人数小于等于概率为0的人数，重新初始化
        if len(临时姓名列表) <= 概率低人数:
            临时姓名列表 = 复制列表(姓名列表)
            临时概率列表 = 复制列表(概率列表)
        # 随机选择并移除已选姓名
        临时姓名变量 = 随机函数(临时姓名列表, 临时概率列表)
        临时姓名列表.pop(临时姓名变量[1])
        临时概率列表.pop(临时姓名变量[1])
        return 临时姓名变量[0]
        
    elif 当前模式 == "与上一个不重复":
        # 模式2：不与上一个重复
        if len(姓名列表) <= 1:
            return 随机函数(姓名列表, 概率列表)[0]
        临时姓名变量 = 随机函数(姓名列表, 概率列表)[0]
        # 循环直到选到不同的姓名
        while 临时姓名变量 == 当前名字:
            临时姓名变量 = 随机函数(姓名列表, 概率列表)[0]
        return 临时姓名变量
        
    else:
        # 模式3：完全随机（可能重复）
        return 随机函数(姓名列表, 概率列表)[0]

def 更新():
    """更新显示的姓名"""
    global 当前名字
    当前名字 =随机()
    名字标签.config(text=当前名字, font=名字字体)
    更新按钮.config(text="换一个")

def 查看概率():
    """显示所有姓名的概率"""
    global 正在查看概率
    正在查看概率 = True
    
    # 使用列表推导和join优化字符串构建
    概率信息列表 = [f"{姓名列表[i]}:{概率列表[i]}" for i in range(len(概率列表))]
    临时变量 = "\n".join(
        "  ".join(概率信息列表[i:i+5]) 
        for i in range(0, len(概率信息列表), 5)
    )
    
    名字标签.config(text=临时变量, font=[名字字体[0], 12])
    更新按钮.pack_forget()  # 隐藏更新按钮
    查看概率按钮.config(text="返回", command=返回)

def 返回():
    """从概率查看界面返回主界面"""
    global 当前名字, 正在查看概率
    正在查看概率 = False
    # 恢复主界面显示
    名字标签.config(text=当前名字, font=名字字体)
    名字标签.pack_forget()
    查看概率按钮.pack_forget()
    隐藏按钮.place(relx=1, rely=0.5, anchor="e")
    标签.pack()
    名字标签.pack()
    更新按钮.pack()
    查看概率按钮.place(relx=0.5, rely=1, anchor="s")
    查看概率按钮.config(text="查看概率", command=查看概率)
    
def 隐藏(动画=True):
    """切换窗口隐藏/显示状态"""
    global 隐藏状态, 正在查看概率, 展开窗口宽度, 展开窗口高度, 窗口顶点x, 窗口顶点y, 收缩帧率, 窗口化状态
    if 隐藏状态:
        # 显示完整窗口
        隐藏状态 = False
        隐藏按钮.place_forget()
        标签.place_forget()
        更多.place_forget()
        标签.pack()
        隐藏按钮.place(relx=1, rely=0.5, anchor="e")
        更多.place(relx=1, rely=1, anchor="se")
        切换模式按钮.place(relx=0, rely=1, anchor="sw")
        名字标签.pack()
        if not 正在查看概率:
            更新按钮.pack()
        查看概率按钮.pack()
        标签.config(text="随机点名", font=("黑体", 8))
        隐藏按钮.config(text="<")
        if 动画:
            for 帧数 in range(收缩帧率):
                窗口.geometry(f"{int((展开窗口宽度 - 收起窗口宽度) / 收缩帧率 * 帧数 + 收起窗口宽度)}x{int((展开窗口高度 - 收起窗口高度) / 收缩帧率 * 帧数 + 收起窗口高度)}+{int(窗口顶点x)}+{int(窗口顶点y)}")
                窗口.update()
        窗口.geometry(f"{展开窗口宽度}x{展开窗口高度}+{int(窗口顶点x)}+{int(窗口顶点y)}")
        显示所有初始组件()
    else:
        # 隐藏为小窗口
        隐藏状态 = True
        if 动画 and not 窗口化状态:
            for 帧数 in range(收缩帧率):
                窗口.geometry(f"{int(展开窗口宽度 - (展开窗口宽度 - 收起窗口宽度) / 收缩帧率 * 帧数)}x{int(展开窗口高度 - (展开窗口高度 - 收起窗口高度) / 收缩帧率 * 帧数)}+{int(窗口顶点x)}+{int(窗口顶点y)}")
                窗口.update()
            窗口.geometry(f"{收起窗口宽度}x{收起窗口高度}+{int(窗口顶点x)}+{int(窗口顶点y)}")
            隐藏所有组件()
            标签.place(relx=0, rely=0.5, anchor="w")
            标签.config(text="随机\n\n\n\n点名", font=("黑体", 8))
            隐藏按钮.place(relx=1, rely=0.5, anchor="e")
            隐藏按钮.config(text="  >")
        elif 窗口化状态:
            关闭动画()
            窗口.geometry(f"{收起窗口宽度}x{收起窗口高度}+{int(窗口顶点x)}+{int(窗口顶点y)}")
            窗口化(change=True)
            隐藏所有组件()
            标签.place(relx=0, rely=0.5, anchor="w")
            标签.config(text="随机\n\n\n\n点名", font=("黑体", 8))
            隐藏按钮.place(relx=1, rely=0.5, anchor="e")
            隐藏按钮.config(text="  >")
            开启动画()
        else:
            窗口.geometry(f"{收起窗口宽度}x{收起窗口高度}+{int(窗口顶点x)}+{int(窗口顶点y)}")
            隐藏所有组件()
            标签.place(relx=0, rely=0.5, anchor="w")
            标签.config(text="随机\n\n\n\n点名", font=("黑体", 8))
            隐藏按钮.place(relx=1, rely=0.5, anchor="e")
            隐藏按钮.config(text="  >")

def 更多选项显示():
    """显示更多选项界面"""
    隐藏所有组件()
    if not 窗口化状态:
        关闭按钮.place(relx=1, rely=0, anchor="ne")
    重启按钮.place(relx=1, rely=1, anchor="se")
    选择文件按钮.place(x=10, y=10, anchor="nw")
    查看概率按钮.place(relx=0.5, rely=1, anchor="s")
    字体大小文本.place(x=10, y=50, anchor="nw")
    字体大小滑块.place(x=85, y=50, anchor="nw")
    窗口化按钮.place(relx=0, rely=1, anchor="sw")
    文件路径文本.place(x=100, y=10, anchor="nw")
    查看概率按钮.config(text="返回", command=初始化所有组件)

def 切换模式():
    """切换随机模式"""
    global 当前模式
    if 当前模式 == "每一轮不重复(推荐)":
        当前模式 = "与上一个不重复"
    elif 当前模式 == "与上一个不重复":
        当前模式 = "仅随机(稳定)"
    else:
        当前模式 = "每一轮不重复(推荐)"
    切换模式按钮.config(text=f"当前模式:{当前模式}")

def 更新字体大小(value=50):
    """更新字体大小"""
    global 名字字体
    名字字体[1] = int(value)
    名字标签.config(font=名字字体)
    字体大小文本.config(text=f"字体大小:{value}")

def 窗口化(change=None):
    """切换窗口化状态"""
    global 窗口化状态, 名字字体
    if not 窗口化状态:
        窗口化状态 = True
        窗口化按钮.config(text="取消窗口化")
        关闭按钮.place_forget()
        窗口.overrideredirect(False)
        字体大小滑块.config(to=400)
    elif change:
        窗口化状态 = False
        窗口化按钮.config(text="窗口化(实验性)")
        # 如果窗口处于全屏状态，先恢复正常状态
        if 窗口.state() == 'zoomed':
            窗口.state('normal')
        # 先设置无边框，再设置位置和大小
        窗口.overrideredirect(True)
        字体大小滑块.config(to=85)
        if 名字字体[1] > 85:
            更新字体大小(value=85)
    else:
        窗口化状态 = False
        窗口化按钮.config(text="窗口化(实验性)")
        关闭按钮.place(relx=1, rely=0, anchor="ne")
        # 如果窗口处于全屏状态，先恢复正常状态
        if 窗口.state() == 'zoomed':
            窗口.state('normal')
        # 先设置无边框，再设置位置和大小
        窗口.overrideredirect(True)
        窗口.geometry(f"{展开窗口宽度}x{展开窗口高度}+{int(窗口顶点x)}+{int(窗口顶点y)}")
        字体大小滑块.config(to=85)
        if 名字字体[1] > 85:
            更新字体大小(value=85)

# 程序初始化
初始化()

# 创建主窗口
窗口 = t.Tk()
窗口.title("随机点名")
窗口.geometry(f"{展开窗口宽度}x{展开窗口高度}+{int(窗口顶点x)}+{int(窗口顶点y)}")  # 设置窗口大小和初始位置
窗口.overrideredirect(True)     # 无边框窗口
窗口.wm_attributes("-topmost", True)  # 窗口置顶

# 创建界面组件
标签 = t.Label(窗口, text="随机点名")
y坐标 = 0
鼠标按下位置 = (0, 0)  # 初始化鼠标按下位置

# 绑定鼠标事件
标签.bind("<Button-1>", 鼠标按下)  # 鼠标按下事件
标签.bind("<B1-Motion>", 移动)    # 鼠标拖动事件

名字字体 = ["宋体", 50] # 设置字体

# 定义组件
关闭按钮 = t.Button(窗口, text="x", command=关闭)
当前模式 = "每一轮不重复(推荐)"
切换模式按钮 = t.Button(窗口, text=f"当前模式:{当前模式}", command=切换模式)
重启按钮 = t.Button(窗口, text="重启", command=重启)
更多 = t.Button(窗口, text="更多...", command=更多选项显示)
隐藏状态 = False
隐藏按钮 = t.Button(窗口, text="<", command=隐藏)
名字标签 = t.Label(窗口, text=当前名字, font=名字字体)
更新按钮 = t.Button(窗口, text="按钮", command=更新)
正在查看概率 = False
查看概率按钮 = t.Button(窗口, text="查看概率", command=查看概率)
当前文件路径 = os.path.abspath("姓名及概率.txt") if 当前名字 != "未找到文件" else "无"
选择文件按钮 = t.Button(窗口, text="选择数据文件", command=选择文件)
文件路径文本 = t.Label(窗口, text=f"当前文件路径:{当前文件路径}")
字体大小文本 = t.Label(窗口, text=f"字体大小:{名字字体[1]}")
字体大小滑块 = t.Scale(窗口, from_=10, to=85, orient="horizontal", command=更新字体大小, showvalue=False)
字体大小滑块.set(名字字体[1])  # 设置滑块初始值
窗口化状态 = False
窗口化按钮 = t.Button(窗口, text="窗口化(实验性)", command=窗口化)


# 显示初始界面
显示所有初始组件()

if __name__ == "__main__":
    # 收起窗口
    隐藏(动画=False)

    # 执行动画
    开启动画()

    # 启动主循环
    窗口.mainloop()