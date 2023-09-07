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
        return i - j
    else:
        return -1


def getNext(P):
    nextArray = [0 for i in range(len(P))]
    nextArray[0] = -1
    j = 0
    k = -1
    while j < len(P) - 1:
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


if __name__ == "__main__":
    A = "ABACBCDHI"
    P = "CD"
    res = KMP(A, P) + 1
    print(A)
    print(P)
    print("第", res, "位开始相同")
