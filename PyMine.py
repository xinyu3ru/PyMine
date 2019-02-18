#!/bin/env python
# -*- coding: utf-8 -*-

"""
依赖 pywin32  win32gui PIL
需要自己下载扫雷，http://www.saolei.net/BBS/，右上角，下载
打开扫雷，然后运行本脚本，全自动扫雷
Win 键强制退出脚本

"""
import random
import win32api
import win32gui
import win32con
import sys
import os
from PIL import ImageGrab

# 查找扫雷游戏窗口
class_name = "TMain"
title_name = "Minesweeper Arbiter "
hwnd = win32gui.FindWindow(class_name, title_name)

# 窗口坐标
left = 0
top = 0
right = 0
bottom = 0

if hwnd:
    print("找到扫雷软件窗口")
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    # win32gui.SetForegroundWindow(hwnd)
    print("窗口坐标为：")
    print(str(left)+' '+str(right)+' '+str(top)+' '+str(bottom))
else:
    print("没有找到扫雷窗口，请先打开扫雷，然后运行本脚本！")

# 雷区坐标，扫雷软件相对比较稳定，雷区位置不变
left += 15
top += 101
right -= 15
bottom -= 42

# 截图雷区图像
rect = (left, top, right, bottom)
img = ImageGrab.grab().crop(rect)

# 雷区每个方块16*16
block_width, block_height = 16, 16
# 计算横向有blocks_x个方块
blocks_x = int((right - left) / block_width)
# 计算纵向有blocks_y个方块
blocks_y = int((bottom - top) / block_height)

# 各种旗帜颜色，多点判断
rgba_ed = [(225, (192, 192, 192)), (31, (128, 128, 128))]
rgba_red = [(54, (255, 255, 255)), (17, (255, 0, 0)), (109, (192, 192, 192)), (54, (128, 128, 128)), (22, (0, 0, 0))]
rgba_0 = [(54, (255, 255, 255)), (148, (192, 192, 192)), (54, (128, 128, 128))]
rgba_1 = [(185, (192, 192, 192)), (31, (128, 128, 128)), (40, (0, 0, 255))]
rgba_2 = [(160, (192, 192, 192)), (31, (128, 128, 128)), (65, (0, 128, 0))]
rgba_3 = [(62, (255, 0, 0)), (163, (192, 192, 192)), (31, (128, 128, 128))]
rgba_4 = [(169, (192, 192, 192)), (31, (128, 128, 128)), (56, (0, 0, 128))]
rgba_5 = [(70, (128, 0, 0)), (155, (192, 192, 192)), (31, (128, 128, 128))]
rgba_6 = [(153, (192, 192, 192)), (31, (128, 128, 128)), (72, (0, 128, 128))]
rgba_8 = [(149, (192, 192, 192)), (107, (128, 128, 128))]

rgba_boom = [(4, (255, 255, 255)), (144, (192, 192, 192)), (31, (128, 128, 128)), (77, (0, 0, 0))]
rgba_boom_red = [(4, (255, 255, 255)), (144, (255, 0, 0)), (31, (128, 128, 128)), (77, (0, 0, 0))]
# 数字1-8
# 0 未被打开
# -1 被打开 空白
# -4 红旗

mine_map = [[0 for i in range(blocks_x)] for i in range(blocks_y)]

# 用于统计是否需要随机点击
random_click = 0
# 是否游戏结束
game_over = 0


# 扫描雷区图像,雷区切块并且判断该块雷的数量
def show_mine_map():
    pic = ImageGrab.grab().crop(rect)
    for y in range(blocks_y):
        for x in range(blocks_x):
            this_image = pic.crop((x * block_width, y * block_height, (x + 1) * block_width, (y + 1) * block_height))
            if this_image.getcolors() == rgba_0:
                mine_map[y][x] = 0
            elif this_image.getcolors() == rgba_1:
                mine_map[y][x] = 1
            elif this_image.getcolors() == rgba_2:
                mine_map[y][x] = 2
            elif this_image.getcolors() == rgba_3:
                mine_map[y][x] = 3
            elif this_image.getcolors() == rgba_4:
                mine_map[y][x] = 4
            elif this_image.getcolors() == rgba_5:
                mine_map[y][x] = 5
            elif this_image.getcolors() == rgba_6:
                mine_map[y][x] = 6
            elif this_image.getcolors() == rgba_8:
                mine_map[y][x] = 8
            elif this_image.getcolors() == rgba_ed:
                mine_map[y][x] = -1
            elif this_image.getcolors() == rgba_red:
                mine_map[y][x] = -4
            elif this_image.getcolors() == rgba_boom or this_image.getcolors() == rgba_boom_red:
                global game_over
                game_over = 1
                break
            else:
                print("无法识别图像")
                print("坐标" + str((y, x)) + "颜色")
                print(this_image.getcolors())
                sys.exit(0)


# 插红旗
def red_flag():
    show_mine_map()
    for y in range(blocks_y):
        for x in range(blocks_x):
            if 1 <= mine_map[y][x] <= 5:
                boom_number = mine_map[y][x]
                block_white = 0
                block_qi = 0
                for yy in range(y-1, y+2):
                    for xx in range(x-1, x+2):
                        if 0 <= yy and 0 <= xx and yy < blocks_y and xx < blocks_x:
                            if not (yy == y and xx == x):
                                if mine_map[yy][xx] == 0:
                                    block_white += 1
                                elif mine_map[yy][xx] == -4:
                                    block_qi += 1
                if boom_number == block_white + block_qi:
                    for yy in range(y - 1, y + 2):
                        for xx in range(x - 1, x + 2):
                            if 0 <= yy and 0 <= xx and yy < blocks_y and xx < blocks_x:
                                if not (yy == y and xx == x):
                                    if mine_map[yy][xx] == 0:
                                        os.system('cls')
                                        print("\n\n全自动扫雷中\n\n按 Win 键强制退出\n插红旗" + str(yy + 1) + ',' + str(xx + 1))
                                        win32api.SetCursorPos([left+xx*block_width, top+yy*block_height])
                                        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
                                        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
                                        show_mine_map()


# 点击空白
def dig():
    show_mine_map()
    is_random_click = 0
    for y in range(blocks_y):
        for x in range(blocks_x):
            if 1 <= mine_map[y][x] <= 5:
                boom_number = mine_map[y][x]
                block_white = 0
                block_qi = 0
                for yy in range(y - 1, y + 2):
                    for xx in range(x - 1, x + 2):
                        if 0 <= yy and 0 <= xx and yy < blocks_y and xx < blocks_x:
                            if not (yy == y and xx == x):
                                if mine_map[yy][xx] == 0:
                                    block_white += 1
                                elif mine_map[yy][xx] == -4:
                                    block_qi += 1
                if boom_number == block_qi and block_white > 0:
                    for yy in range(y - 1, y + 2):
                        for xx in range(x - 1, x + 2):
                            if 0 <= yy and 0 <= xx and yy < blocks_y and xx < blocks_x:
                                if not(yy == y and xx == x):
                                    if mine_map[yy][xx] == 0:
                                        print("点开" + str(yy + 1) + ',' + str(xx + 1))
                                        win32api.SetCursorPos([left + xx * block_width, top + yy * block_height])
                                        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                                        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                                        is_random_click = 1
    if is_random_click == 0:
        random_click()


# 随机点击
def random_click():
    fl = 1
    while fl:
        random_x = random.randint(0, blocks_x - 1)
        random_y = random.randint(0, blocks_y - 1)
        if not mine_map[random_y][random_x]:
            print("随机点击" + str(random_y + 1) + ',' + str(random_x + 1))
            win32api.SetCursorPos([left + random_x * block_width, top + random_y * block_height])
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            fl = 0


def main():
    win32api.SetCursorPos([left, top])
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    show_mine_map()
    global game_over
    while 1:
        if not game_over:
            red_flag()
            red_flag()
            dig()
        else:
            game_over = 0
            win32api.keybd_event(113, 0, 0, 0)
            win32api.SetCursorPos([left, top])
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            show_mine_map()


if __name__ == "__main__":
    main()
