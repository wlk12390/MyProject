from numpy import *
import numpy as np
import random
m=12 #原始报文数量
n=36 #编码报文数量
# k=5
#初始化
data = np.zeros(m) #原始报文
datadecode = np.zeros(m) #解码报文
number=[]
number1=[]
for i in range(0,m):
    data[i] = random.randint(1,50)
for i in range(0, 2**8):
    num = i
    number.append(num)
for i in range(0, n):
    num1 = i
    number1.append(num1)
#print(number)
print(" 原始报文","\n",data)
g = np.zeros((n,m)) #编码系数
y = np.zeros((n,m)) #编码报文
receive = np.zeros((m,1)) #接受报文
greceive = np.zeros((m,m)) #系数矩阵
#print(receive,greceive)
#选取系数
for i in range(0,n):
    for j in range(0,m):
        g[i][j] = random.choice(number)
#编码
for i in range(0,n):
    for j in range(0,m):
        y[i][j] = g[i][j]*data[j]
B=np.array(y)
#随机抓取m个消息进行解码
for i in range(0,m):
    temp = random.choice(number1)
    receive[i] = sum(B[temp])
    for j in range(0, m):
        greceive[i][j]=g[temp][j]
# print(receive)
# print(greceive)
receive=np.array(receive)
greceive=mat(greceive)
#解码
for l in range(10000):
    #判断矩阵是否满秩
    if np.linalg.matrix_rank(greceive)==m:
        # print(greceive.I,np.linalg.det(greceive))
        natadecode = np.transpose(np.dot(greceive.I,receive))
        print(" datadecode","\n",natadecode)
        break
    # 选取系数
    for i in range(0, n):
        for j in range(0, m):
            g[i][j] = random.choice(number)
    # 编码
    for i in range(0, n):
        for j in range(0, m):
            y[i][j] = g[i][j] * data[j]
    B = np.array(y)
    greceive = np.zeros((m, m))
    receive = np.zeros((m, 1))
    # 随机抓取m个消息进行解码
    for i in range(0, m):
        temp = random.choice(number1)
        receive[i] = sum(B[temp])
        for j in range(0, m):
            greceive[i][j] = g[temp][j]
    #print(greceive)
    receive = np.array(receive)
    greceive = mat(greceive)
    # print(l,greceive)
    # print(receive)
    if l>=20:
        print("传输失败")
        break



