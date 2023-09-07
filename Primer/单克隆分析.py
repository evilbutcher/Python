#!/usr/bin/python
# -*- coding: UTF-8 -*-

import pyperclip
import os


def KMP(A, P):  # O(M+N)
    i = 0  # 主串的位置
    j = 0  # 模式串的位置
    nextArray = getNext(P)
    while i < len(A) and j < len(P):
        if j == -1 or A[i] == P[j]:
            i += 1
            j += 1
        else:  # i不回溯
            j = nextArray[j]  # j回到指定位置
    if j == len(P):
        return i-j
    else:
        return -1


def getNext(P):
    nextArray = [0 for i in range(len(P))]
    nextArray[0] = -1
    j = 0
    k = -1
    while j < len(P)-1:
        if k == -1 or P[j] == P[k]:
            j += 1
            k += 1
            if P[j] == P[k]:  # 两个字符相等跳过
                nextArray[j] = nextArray[k]
            else:
                nextArray[j] = k
        else:
            k = nextArray[k]
    return nextArray


if __name__ == '__main__':

    N = input('请输入有几条链')
    N = int(N)
    R = input('请输入想要检测的相同碱基数')
    R = int(R)
    a = 'y'
    x = 1
    dna = []
    print('请拷贝DNA链')
    a = input('y or n')
    dna.append(pyperclip.paste())
    while x < N:
        if (a == 'y'):
            l = os.system("cls")
            print('请继续拷贝DNA链,您刚才拷贝的是', (dna[x-1]))
            a = input('y or n')
            dna.append(pyperclip.paste())
            x = x + 1
    l = os.system("cls")
    print(dna)
    print()
    print()

n = N - 1
g = 0
while g < n:
    h = 0
    while h <= n:
        if h > g:
            print(dna[h])
            print(dna[g])
            w = 0
            while w < len(dna[g])-4:
                str = dna[g][w:w+R]
                A = dna[h]
                O = dna[g]
                P = str
                res = KMP(A, P)+1
                if res >= 1:
                    print('重复片段为', P,)
                    print('从第一行', res, '位开始相同')
                    print()
                    print()
                w = w+1
        h = h+1
    g = g+1

input("Press <Enter>")
