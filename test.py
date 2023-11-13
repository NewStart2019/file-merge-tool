#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import math
from decimal import Decimal
import string


# 设置最大递归层数为1000
sys.setrecursionlimit(1000)


# 从数组中获取最小值
def minValue(list=[]):
    if len(list) == 0:
        return 0
    temp = list[0]
    for i in list:
        if i < temp:
            temp = i
    return temp


# 从数组中获取最大值
# 56, 12, 6, 78, 100   n数字，比较多少次？ n-1
# temp = 78
def maxValue(list=[]):
    if len(list) == 0:
        return 0
    temp = list[0]
    for i in list[1:]:
        if i > temp:
            temp = i
    return temp


# 基础数据类型
# TODO 自己看 set、dict 的 增、改、移除
def baseDataType() -> None:
    print(type(1))
    print(type(3.1415926))
    print(type("3.14"))
    print(type(float("3.14")))
    # 打印集合类型
    print(type([]))  # list
    list = [56, 12, 6, 78, 100]
    list.append(101)
    list.extend([102, 103])
    print(list.pop(0))
    list.remove(103)
    print(list)

    print(type(()))  # tuple

    a1 = {1, 2, 3, 4, 5, 6}
    a2 = {4, 5, 6, 7, 8}
    print(type(a1))  # set
    # 集合合并
    print(a1.union(a2))

    print(type({'name': '龚江红'}))  # dict
    dict1 = {'a': 1, 'b': 2}
    dict2 = {'c': 3, 'd': 4}
    dict2['d'] = "thank"
    # 字典合并
    dict1.update(dict2)
    print(dict1)


# TODO 利用递归计算 等差数列 sum = 1 + 4 + 7 + 10 + …… 计算前1000项的和

# 阶乘函数 3! = 3 * jie(3-1) = 3*2 = 6
#                     2*jie(2-1) 2*1=2
#                         1

# 递归计算阶乘
def jie(v) -> int:
    if v == 1:
        return 1
    return v * jie(v - 1)


# 计算浮点数小数长度
def floatLength(v) -> int:
    if v == None:
        return 0
    return  len(str(v).split(".")[1])

# TODO 3 实现泰勒展开式  2.71828 e^1 = 2.71828 ，优化一下，为什么有0.2以下的误差
# e^x = 1 + 1/1! * x + 1/2! * x^2 + 1/ 3! * x^3 + ……
# 之前函数的问题 1、重复计算阶乘 1！ 2! 3! 4!
#          2、float的小数位数最多为15，那么就不能精确到15为之后，但是15位精确度不够
def calculateE(length=10, x=1) -> Decimal:
    x = Decimal(x)
    one = Decimal("1.0")
    counter = one
    length = Decimal(length)
    result = one
    jie = counter
    while (True):
        temp = Decimal(one / jie) * Decimal(math.pow(x, counter))
        result += temp
        if counter >= length:
            break
        counter += one
        jie = jie * counter
    return result

print("2 => 龚江红不爱学习python！")


"""
龚江红
认真
学
"""
if __name__ == '__main__':
    with open(file="print.txt", mode="a") as f:
        # 输出打印内容到print.txt
        print("1 => 龚江红学习python", "gb2312", sep="=====", end="", file=f)
    list = [56, 12, 6, 78, 100]
    print(minValue(list))
    print(maxValue(list))
    baseDataType()
    a,b = list[0:2]
    print(a,b)
    print(calculateE(10, 1))
