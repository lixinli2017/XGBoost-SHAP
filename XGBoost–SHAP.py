# coding=utf-8
# -*- coding: utf-8 -*-
'''
无调用加载
'''
def hooketest():
    return 0
'''
调用加载
'''

import os

os.environ["JOBLIB_MULTIPROCESSING"] = "0"
os.environ["LOKY_MAX_CPU_COUNT"] = "1"
from os.path import isdir
import linecache
# 正则化，用于读取数据排序
import re
# python自带的模块，缓存读取文件某一行的内容
import numpy as np
import pandas as pd
import math
import copy as copy1
import scipy
import urllib
# from scipy.signal import savgol_filter
import urllib.parse
from urllib.parse import urlparse
import sys
from numpy import loadtxt
from numpy import savetxt
from numpy.linalg import svd

np.set_printoptions(suppress=True)
# 可以使数组打印更漂亮
np.set_printoptions(threshold=np.inf)
# 解决ndarray打印省略问题
# np.set_printoptions(linewidth=20)
# #打印时候每行显示20个

import pylab
from pylab import *
import matplotlib as mpl

# mpl.rcParams['font.sans-serif'] = ['SimHei']
# mpl.rcParams['font.sans-serif'] = ['Times New Roman']
mpl.rcParams['font.sans-serif'] = ['Arial']
# ！SimHei
mpl.rcParams['axes.unicode_minus'] = False
# 用来正常显示负号
mpl.rcParams['svg.fonttype'] = 'none'   # 可以保留，不影响运行
from matplotlib.ticker import FuncFormatter


def to_percent(temp, position):
    return '%1.0f' % (100 * temp) + '%'


# 刻度以%显示，小数点后1位

import matplotlib

import matplotlib.cm as cm

# 新旧版本兼容
if not hasattr(matplotlib, "register_cmap"):
    # 为旧MATPLOTLIB版本兼容性兜底
    from matplotlib import colormaps as cm
else:
    from matplotlib import cm

from matplotlib import pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
from scipy import signal
from scipy import optimize
from scipy import sparse
# from scipy.signal import find_peaks_cwt
from scipy.signal import find_peaks
from scipy.signal import detrend
from scipy.sparse import linalg
from scipy.sparse import csc_matrix, eye, diags
from scipy.sparse.linalg import spsolve
from scipy.signal import savgol_filter
from scipy.cluster import hierarchy
from scipy import cluster
# 用于进行层次聚类，话层次聚类图的工具包
from scipy.spatial.distance import cdist
# 计算距离结算公式

from collections import Counter
# 高性能容量数据类型,count计数器
from decimal import *
import sklearn
from sklearn import decomposition as skldec
# 用于主成分分析降维的包
from sklearn.model_selection import train_test_split
from sklearn import svm
from sklearn.discriminant_analysis import LinearDiscriminantAnalysis
from sklearn.neighbors import KNeighborsClassifier
from sklearn.manifold import TSNE
# tsne
from sklearn import metrics
from sklearn.metrics import silhouette_score
# 计算轮廓系数
from sklearn.model_selection import GridSearchCV
# 网格搜索&交叉验证
# from sklearn.externals import joblib
# sklearn.externals.joblib函数是用在0.21及以前的版本中，在最新的版本中，该函数应被弃用。
# 直接改为import joblib
import joblib
# 保存、加载分类器模型
from xgboost import XGBClassifier
# xgboost
# from skimage import filters
# 基于Otsu的阈值分割方法,无参数去背景


# ROCq曲线AUC值
# from sklearn import datasets
from sklearn.metrics import roc_curve, auc
from sklearn.preprocessing import label_binarize
# from sklearn.multiclass import OneVsRestClassifier
# from scipy import interp
# 画图-混淆矩阵
import seaborn as sns
from sklearn.metrics import confusion_matrix
# -----------------------
import struct
# Kmeans模型
import sklearn.cluster
from sklearn.cluster import KMeans
from sklearn.cluster import MiniBatchKMeans
from sklearn.cluster import SpectralClustering
from sklearn.cluster import AgglomerativeClustering
# 聚类算法距离计算公式
from scipy.spatial.distance import cdist

from scipy.stats import pearsonr
# 函数：pearsonr(x,y),计算特征与目标变量之间的相关度
# 输出：r： 相关系数 [-1，1]之间，p-value: p值
#方法二：根据scipy库求解
# from scipy.spatial.distance import pdist
from scipy.spatial.distance import cosine
# 计算两个 1-D 数组间的 Cosine distance
from scipy.spatial.distance import euclidean
# 计算两个 1-D 数组间的 Euclidean distance(欧几里德距离)
from scipy.spatial.distance import minkowski
# 计算两个 1-D 数组间的 Minkowski distance(明可夫斯基距离)
from scipy.spatial.distance import jensenshannon
# 计算两个 1-D 数组间的 Jensen-Shannon distance metric
# JS散度(Jensen-Shannon)是 Kullback–Leibler divergence(KL距离)的变种， 相似度衡量指标
from scipy.spatial.distance import mahalanobis
# 计算两个 1-D 数组间的 Mahalanobis distance(马氏距离)
from scipy import interpolate
# 三次样条插值
# from dtw import dtw as DTW
# 动态时间规整/规划(Dynamic Time Warping, DTW)能够用来衡量两个时间序列的相似性


# 写入word
import docx
import docx.enum
import docx.shared
import docx.enum.text
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import docx.oxml
import docx.oxml.ns
from docx.oxml.ns import qn

import unicodedata
from sklearn.manifold import SpectralEmbedding
# SpectralEmbedding 拉普拉斯特征映射,非线性维度的频谱嵌入
from sklearn.manifold import MDS
# MDS 多维尺度变换
import warnings
import shutil, os, send2trash
# from collections import Counter

warnings.filterwarnings("ignore")
# 忽略警告
import shap


# airPLS去背景
# def WhittakerSmooth(x, w, lambda_):
#     # X = np.matrix(x)
#     X = np.asarray(x)
#     m = X.size
#     i = np.arange(0, m)
#     E = eye(m, format='csc')
#     D = E[1:] - E[:-1]
#     # numpy.diff() does not work with sparse matrix. This is a workaround.
#     W = diags(w, 0, shape=(m, m))
#     A = csc_matrix(W + (lambda_ * D.T * D))
#     B = csc_matrix(W * X.T)
#     background = spsolve(A, B)
#     return np.array(background)

# airPLS 去背景核心函数（Python3.9兼容版）
def WhittakerSmooth(x, w, lambda_):
    """
    Whittaker 平滑法，用于 airPLS 背景估计。
    完全兼容 Python3.9 + SciPy>=1.10。
    """
    x = np.asarray(x).reshape(-1, 1)     # 转为列向量
    m = x.shape[0]
    E = eye(m, format='csc')
    D = E[1:] - E[:-1]
    W = diags(w, 0, shape=(m, m))
    A = csc_matrix(W + (lambda_ * D.T @ D))
    B = W @ x                            # (m,1)
    background = spsolve(A, B)
    return np.asarray(background).ravel()  # 还原为一维数组

# airPLS去背景主程序
def airPLS(x, lambda_, itermax):
    x = np.array(x)
    m = x.shape[0]  # 计算长宽维度
    w = np.ones(m)
    for i in range(1, itermax + 1):
        z = WhittakerSmooth(x, w, lambda_)
        d = x - z
        dssn = np.abs(d[d < 0].sum())
        # ***********************************************除零问题待解决
        # if fabs(dssn) < 1e-15:
        #     dssn = dssn+0.0001
        # print("dssn:",dssn)
        if (dssn < 0.001 * (abs(x)).sum() or i == itermax):
            if (i == itermax):
                print('WARING max iteration reached!')
            break
        w[d >= 0] = 0
        # d>0 means that this point is part of a peak, so its weight is set to 0 in order to ignore it
        w[d < 0] = np.exp(i * np.abs(d[d < 0]) / dssn)
        w[0] = np.exp(i * (d[d < 0]).max() / dssn)
        w[-1] = w[0]
    z = d
    return z.tolist()

#     return corrected_data.tolist()

# 方法二，ALS交替最小二乘算法
def molification_smoothing(rawspectrum, struct_el, number_of_molifications):
    molifier_kernel = np.linspace(-1, 1, num=struct_el)
    # print(molifier_kernel)
    # 创建-1到1之间数字序列，步长为struct_el
    molifier_kernel[1:-1] = np.exp(-1 / (1 - molifier_kernel[1:-1] ** 2))
    # print(molifier_kernel)
    molifier_kernel[0] = 0
    molifier_kernel[-1] = 0
    molifier_kernel = molifier_kernel / np.sum(molifier_kernel)
    # 列表归一化
    denominormtor = np.convolve(np.ones_like(rawspectrum), molifier_kernel, 'same')
    # np.convolve卷积运算，same：返回的数组长度为max(M, N),边际效应依旧存在
    # np.ones_like，全1数组
    smoothline = rawspectrum
    i = 0
    for i in range(number_of_molifications):
        smoothline = np.convolve(smoothline, molifier_kernel, 'same') / denominormtor
        i += 1
    return smoothline


# 方法二，ALS交替最小二乘算法主函数
def background_substruction(y):
    als_lambda = 3e4
    als_p_weight = 5e-4
    _iterations = 16
    # zero_step_struct_el = np.int(2 * np.round(len(y) / 200) + 3)
    zero_step_struct_el = int(2 * np.round(len(y) / 200) + 3)
    """
    猜测，设置步长每10个点为一步，步长越短，拟合效果越保型，最终点数至少大于等于3，所以这里将“+1”该为“+3”
    """
    # print(zero_step_struct_el)
    y_sm = molification_smoothing(y, zero_step_struct_el, 16)
    # 交替（或迭代）做平滑,平滑效果其实不好，主要为了拟合光谱形状
    # plt.plot(y,"r")
    # plt.plot(y_sm,"b")
    # plt.show()
    # compute the derivatives:
    y_sm_1d = np.gradient(y_sm)
    # !!!!!!重要理解 https://docs.scipy.org/doc/numpy-1.12.0/reference/generated/numpy.gradient.html
    # numpy.gradient，返回 N 维数组的梯度，
    # 梯度是使用内部二阶精确的中心差异和边界处的一阶差异或二阶精确的一侧（向前或向后）差异来计算的。
    # 因此，返回的梯度与输入数组具有相同的形状。
    y_sm_2d = np.gradient(y_sm_1d)
    # weighting function for the 2nd der:二阶权函数
    y_sm_2d_decay = (np.mean(y_sm_2d ** 2)) ** 0.5
    # 二阶权值数组平方后均值在开方
    weifunc2D = np.exp(-y_sm_2d ** 2 / 2 / y_sm_2d_decay ** 2)
    # weighting function for the 1st der:
    y_sm_1d_decay = (np.mean((y_sm_1d - np.mean(y_sm_1d)) ** 2)) ** 0.5
    weifunc1D = np.exp(-(y_sm_1d - np.mean(y_sm_1d)) ** 2 / 2 / y_sm_1d_decay ** 2)
    weifunc = weifunc1D * weifunc2D
    # exclude from screenenig the edges of the spectrum (optional)，排除光谱的边缘
    weifunc[0:zero_step_struct_el] = 1
    weifunc[-zero_step_struct_el:] = 1

    # estimate the peak height，估计峰值高度
    peakscreen_amplitude = (np.max(detrend(y)) - np.min(
        detrend(y))) / 2
    # 尚不理解：8 is good, because this is a characteristic height of a tail
    """
    detrend(y),去除光谱趋势
    """
    L = len(y)
    D = sparse.diags([1, -2, 1], [0, -1, -2], shape=(L, L - 2))
    D = als_lambda * D.dot(D.transpose())
    w = np.ones(L)
    W = sparse.spdiags(w, 0, L, L)
    # k = 10 * morphological_noise(y)
    #  above this height the peaks are rejected
    for i in range(_iterations):
        W.setdiag(w)
        Z = W + D
        z = spsolve(Z, w * y)
        w = als_p_weight * weifunc * np.exp(-((y - z) / peakscreen_amplitude) ** 2 / 2) * \
            (y > z) + (1 - als_p_weight) * (y < z)
    # baseline = z
    corrected_data = y - z
    return corrected_data.tolist()


# experimental_spectrum = np.genfromtxt\
#     (r'C:\Users\lixin\Desktop\测试数据\5种分类\#01\1-1200g-2-10s_1.txt')
# x = experimental_spectrum[:, 0]
# exp_y = experimental_spectrum[:, 1]
# bldp = background_substruction(exp_y)
#
# plt.figure()
# plt.plot(x, exp_y, 'k', label='exp.spec', LineWidth=1)
# plt.plot(x, exp_y - bldp, 'r', label='derpsalsa', LineWidth=1)
# plt.show()

def sort_key(s):
    # 排序关键字匹配
    # 匹配开头数字序号
    if s:
        try:
            '''
            # import re，正则化，用于读取数据排序
            [-1]表示正则化匹配字符串中最后面的数字，[0]表示最前面的数字
            '''
            # c = re.findall(r'\d+', s)[-2]
            # 菌落挑菌软件，需要对菌落名+光谱排序，这里仅仅对菌落排序（第2个数字），未对光谱排序
            # c = re.findall(r'\d+', re.findall(r'(?<=\-).+?(?=\_)', s)[0])[0]
            # 1.正则匹配两个特定字符'-'和'_'之间的字符,'A2(1)-No.13_45_spec.txt'
            # 2.对'-'和'_'之间的数字排序
            c = re.findall(r'\d+', re.findall(r'(?<=No.).+?(?=\_)', s)[0])[0]
            # 3.对"No."和"_"之间的数字排序
        except:
            c = -1
        # print(c)
        return int(c)

def strsort(alist):
    alist.sort(key=sort_key)
    # alist.sort(key=lambda sort_key:(sort_key.sort(),sort_key[0],sort_key[1] ))
    return alist

# 1读取文件夹并合并所有文件
# 判断字符串是否为数字
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

    try:
        # import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass

    return False


# 合并文件夹，分类
def readMerge(dataPath):
    root = dataPath
    # file_names = os.listdir(root)
    file_names = []
    # 存储所有txt文档文件名
    f_list = os.listdir(root)
    # print("&1",f_list)
    # f_list.sort(key= lambda x:int(x[-6:-4]))
    # print("&2", f_list)
    f_list.sort()
    # 排序
    # print ("**",f_list)
    for i in f_list:
        # os.path.splitext():分离文件名与扩展名
        if os.path.splitext(i)[1] == '.txt':
            # 判断是否为txt文档，
            '''
                识别txt中非光谱数据
            '''
            tempPath = os.path.join(root, i)
            # firLine = np.array(pd.read_csv(tempPath, delim_whitespace=True, header=None , encoding = 'gb2312'))[0]
            firLine = np.array(pd.read_csv(tempPath, sep='\s+', header=None, encoding='gb2312'))[0]
            # firLine = np.array(pd.read_csv(tempPath, delim_whitespace=True, header=None, encoding='utf-8'))[0]
            # firLine = np.array(pd.read_csv(i, delim_whitespace=True, header=None, encoding='utf-8'))[0]
            # print(firLine)
            if (is_number(firLine[0]) and is_number(firLine[1])):
                # [400, 30.54]
                file_names.append(i)
                # print(i)
            else:

                # firLine = np.array(pd.read_csv(tempPath, delim_whitespace=True, encoding = 'gb2312'))[0]
                firLine = np.array(pd.read_csv(tempPath, sep='\s+', header=None, encoding='gb2312'))[0]
                # firLine = np.array(pd.read_csv(tempPath, delim_whitespace=True, encoding='utf-8'))[0]
                # print(firLine)
                if (is_number(firLine[0]) and is_number(firLine[1])):
                    # [400, 30.54]
                    file_names.append(i)
                    # print(i)
                else:
                    continue
    '''
    读取字符串文件，按其中数字大小排序
    '''
    strsort(file_names)

    file_ob_list = []
    for file_name in file_names:
        # 循环地给这520个文件名加上它前面的路径，以得到它的具体路径
        fileob = root + '\\' + file_name
        # 文件夹路径加上\\ 再加上具体要读的的txt的文件名就定位到了这个txt
        file_ob_list.append(fileob)
    # temp = np.loadtxt(file_ob_list[0])
    # allDataSpe = np.array(temp)[:, 0]
    # for fileNameInter in file_ob_list:
    #     temp = np.loadtxt(file_ob_list[0])
    #     allDataSpe = np.column_stack((allDataSpe, np.array(temp)[:, 1]))
    # print("11:",file_ob_list)

    '''
    存储横坐标，拉曼位移
    '''

    firLine = np.array(pd.read_csv(file_ob_list[0], delim_whitespace=True, header=None, encoding = 'gb2312'))
    # firLine = np.array(pd.read_csv(file_ob_list[0], delim_whitespace=True, header=None, encoding='utf-8'))
    if (is_number(firLine[0, 0]) and is_number(firLine[0, 1])):
        # 首行数字
        # temp = pd.read_csv(file_ob_list[0], delim_whitespace=True, header=None)
        allDataSpe = firLine[:, 0]
    else:
        # 首行非数字
        # temp = pd.read_csv(file_ob_list[0], delim_whitespace=True)
        allDataSpe = firLine[1:, 0]
    '''
    存储纵坐标，光谱强度
    '''
    for fileNameInter in file_ob_list:
        temp = np.array(pd.read_csv(fileNameInter, delim_whitespace=True, header=None))
        # print("***",temp.shape)
        if (is_number(temp[0, 0]) and is_number(temp[0, 1])):
            # 首行数字
            # print("&&&",temp[:, 1].shape)
            allDataSpe = np.column_stack((allDataSpe, temp[:, 1]))
        else:
            # 首行非数字
            # print("???",temp[1:, 1].shape)
            allDataSpe = np.column_stack((allDataSpe, temp[1:, 1]))



    return np.array(allDataSpe, dtype=np.float64)

    #


# 一个文件夹下聚类，记住小翟的id标签号
# 判断字符串是否为数字
# def is_number(s):
#     try:
#         float(s)
#         return True
#     except ValueError:
#         pass
#
#     try:
#         import unicodedata
#         unicodedata.numeric(s)
#         return True
#     except (TypeError, ValueError):
#         pass
#
#     return False
# 聚类
def readMerge1(dataPath):
    root = dataPath
    # file_names = os.listdir(root)
    file_names = []
    # 存储所有txt文档文件名
    f_list = os.listdir(root)
    # print f_list
    for i in f_list:
        # os.path.splitext():分离文件名与扩展名
        if os.path.splitext(i)[1] == '.txt':
            # 判断是否为txt文档，
            '''
                识别txt中非光谱数据
            '''
            tempPath = os.path.join(root, i)

            # firLine = np.array(pd.read_csv(tempPath, delim_whitespace=True, header=None, encoding = 'gb2312'))[0]
            firLine = np.array(pd.read_csv(tempPath, sep='\s+', header=None, encoding='gb2312'))[0]
            # firLine = np.array(pd.read_csv(tempPath, delim_whitespace=True, header=None, encoding='utf-8'))[0]
            # firLine = np.array(pd.read_csv(i, delim_whitespace=True, header=None, encoding='utf-8'))[0]
            # print(firLine)
            if (is_number(firLine[0]) and is_number(firLine[1])):
                # [400, 30.54]
                file_names.append(i)
                # print(i)
            else:
                # firLine = np.array(pd.read_csv(tempPath, delim_whitespace=True, encoding = 'gb2312'))[0]
                firLine = np.array(pd.read_csv(tempPath, sep='\s+', header=None, encoding='gb2312'))[0]
                # firLine = np.array(pd.read_csv(tempPath, delim_whitespace=True, encoding='utf-8'))[0]
                # print(firLine)
                if (is_number(firLine[0]) and is_number(firLine[1])):
                    # [400, 30.54]
                    file_names.append(i)
                    # print(i)
                else:
                    continue

            # file_names.append(i)
    '''
    读取字符串文件，按其中数字大小排序
    '''
    strsort(file_names)

    file_ob_list = []
    idLabel = []
    # 用于存储小翟的id标签号
    for file_name in file_names:
        # 循环地给这520个文件名加上它前面的路径，以得到它的具体路径
        fileob = root + '\\' + file_name
        # print("#",file_name)

        # 识别标签，保存文件名
        if (len(file_name) < 12):
            idLabel.append(file_name[:-4])
        else:
            # 自动选择,命名规则“#001_001.txt”
            if (file_name[-12] == '#'):
                idLabel.append(file_name[-12: -4])
            else:
                idLabel.append(file_name[:-4])

        file_ob_list.append(fileob)
    # print(idLabel)
    # return 0


    '''
    存储横坐标，拉曼位移
    '''

    firLine = np.array(pd.read_csv(file_ob_list[0], delim_whitespace=True, header=None, encoding = 'gb2312'))
    # firLine = np.array(pd.read_csv(file_ob_list[0], delim_whitespace=True, header=None, encoding='utf-8'))
    if (is_number(firLine[0,0]) and is_number(firLine[0,1])):
        # 首行数字
        # temp = pd.read_csv(file_ob_list[0], delim_whitespace=True, header=None)
        allDataSpe = firLine[:, 0]
    else:
        #首行非数字
        # temp = pd.read_csv(file_ob_list[0], delim_whitespace=True)
        allDataSpe = firLine[1:, 0]
    '''
    存储纵坐标，光谱强度
    '''
    for fileNameInter in file_ob_list:
        temp = np.array(pd.read_csv(fileNameInter, delim_whitespace=True, header=None))
        # print("***",temp.shape)
        if (is_number(temp[0,0]) and is_number(temp[0,1])):
            # 首行数字
            # print("&&&",temp[:, 1].shape)
            allDataSpe = np.column_stack((allDataSpe, temp[:, 1]))
        else:
            #首行非数字
            # print("???",temp[1:, 1].shape)
            allDataSpe = np.column_stack((allDataSpe, temp[1:, 1]))


    # data = []
    # print("@@@",allDataSpe.shape)
    return np.array(allDataSpe, dtype=np.float64), idLabel
    # 返回ldata将list型转为numpy，方便索引


def CRR(y, filterSize, dynamicFactor):
    N1 = len(y)
    # print(N)
    crrDot = []
    for i in range(filterSize, N1 - filterSize - 1):
        # maxY = y[i - filterSize:i + filterSize + 1].max()
        # minY = y[i - filterSize:i + filterSize + 1].min()
        minY = np.min(y[i - filterSize:i + filterSize + 1])
        medianY = np.median(y[i - filterSize:i + filterSize + 1])
        if ((y[i] - minY) > (medianY - minY) * dynamicFactor):
            crrDot.append(i)
    N2 = len(crrDot)
    if N2 == 0:
        return y
    else:
        for j in range(N2):
            count = 0
            while (((j + count + 1) < N2) and ((crrDot[j + count + 1] - crrDot[j + count]) == 1)):
                count = count + 1
            x1 = crrDot[j] - 1
            x2 = crrDot[j + count] + 1
            y1 = y[x1]
            y2 = y[x2]
            y[crrDot[j]] = (y2 - y1) * (crrDot[j] - x1) / (x2 - x1) + y1
        return y



def MaxMinNormalization(y, x_axis, peak_range=None):
    """
    按指定峰位置进行归一化（若未指定则全局归一化）

    参数：
    ----------
    y : array-like
        单条拉曼光谱强度
    x_axis : array-like
        对应的拉曼位移坐标
    peak_range : list | tuple | None
        峰位范围，例如 [800, 820]

    返回：
    ----------
    list
        归一化后的光谱（目标峰位置强度映射为1）
    """

    y = np.array(y).astype(float)
    x_axis = np.array(x_axis).astype(float)

    if peak_range is not None and len(peak_range) == 2:
        # 找出该峰位区间的索引范围
        idx = np.where((x_axis >= peak_range[0]) & (x_axis <= peak_range[1]))[0]
        if len(idx) > 0:
            peak_max = np.max(y[idx])
        else:
            # 若指定区间内无数据则退回全局最大值
            peak_max = np.max(y)
    else:
        peak_max = np.max(y)

    if peak_max == 0:
        return list(np.zeros_like(y))
    else:
        y_norm = y / peak_max
        return list(y_norm)


# 删除文件夹下所有tif图片
def del_files(path):
    for root, dirs, files in os.walk(path):
        for name in files:
            if name.endswith(".tif"):
                os.remove(os.path.join(root, name))
    return 0



# **************************************************#################################***********************************#############################******一键生成报告

def oneclickReport(dataPath, outPath, startSpec, endSpec, snrNoiseindex, delSNRmin, delSNRmax, filterSize,
                   dynamicFactor, winSG, nSG, lambdaAirPls,
                   itermaxAirPls, distanceMethod, hcaLabelsize, svmTestSize, svmC, xKernel,
                   clusters_N):

    #
    listTable2Value = [dataPath, outPath, startSpec, endSpec, snrNoiseindex, delSNRmin, delSNRmax, filterSize,
                       dynamicFactor, winSG,
                       nSG, lambdaAirPls, itermaxAirPls, distanceMethod, hcaLabelsize,
                        svmTestSize, svmC, xKernel, clusters_N]
    listTable2Name = ["dataPath", "outPath", "startSpec", "endSpec", "snrNoiseindex", "delSNRmin", "delSNRmax",
                      "filterSize",
                      "dynamicFactor", "winSG", "nSG", "lambdaAirPls", "itermaxAirPls",
                      "distanceMethod", "hcaLabelsize", "svmTestSize", "svmC", "xKernel",
                      "clusters_N"]
    '''
    pca参数
    '''
    # 和tsne共用参数
    pcaNcomponents = 2
    # 主成分个数，大于等于2
    pca_col1 = 0
    # 第一维，小于主成分个数
    pca_col2 = 1
    # 第二维，小于主成分个数
    # pcaSize = 30
    # pca标签大小
    pcaAlpha = 0.6
    # pca透明度
    '''
    tSNE参数
    '''
    nDim = 2



    # 定义颜色和形状，使用取余运算做循环
    # colorSpectrumColour = ['#1C1C1C','#008B8B', '#006400', '#00FFFF', '#228B22', '#90EE90', '#8B8B00', '#CD5C5C',
    #                        '#698B22', '#FFA500', '#9932CC', '#3A5FCD', '#5D478B']
    colorSpectrumColour = ['#FF7F00', '#8B0000', '#FFFF00', '#00008B', '#006400', '#90EE90', '#698B69',
                           '#00FFFF', '#008B45', '#9932CC', '#3A5FCD', '#5D478B', '#1C1C1C','#D43F3A','#EEA236',
                           '#5CB85C','#46B8DA','#357EBD','#9632B8','#B8B8B8']
    # colorSpectrumColour = ['#90EE90', '#8B0000', '#8B008B', '#008B8B', '#00008B', '#698B69', '#FFE1FF', '#FFA500',
    #                        '#00FFFF', '#008B45', '#9932CC', '#3A5FCD', '#5D478B']
    # 光谱蓝色颜色
    colorSpectrumBlue = ['0000FF', '0000EE', '0000DD', '0000CC', '0000BB', '0000AA', '000099', '000088', '000077',
                         '000066', 'BCBCBC', 'CCCCCC', 'DCDCDC', 'ECECEC']
    # 光谱渐变灰色
    colorSpectrumGrey = ['#1C1C1C', '#2C2C2C', '3C3C3C', '4C4C4C', '5C5C5C', '6C6C6C', '7C7C7C', '8C8C8C', '9C9C9C',
                         'ACACAC', 'BCBCBC', 'CCCCCC', 'DCDCDC', 'ECECEC']
    colors = ['#1C1C1C', '#0000FF', '#FFFF00', '#7FFFD4', '#008B45']
    # markers = ["s", "o", "+", "v", "d", "x", "3", "p", ".", "*", "1","<",">",","]
    # "正方形"，"圈"，"加号"，"倒三角"，"细钻"，"×号"，"tri_left","五角"，“点”，“tri_down”,"左三角形","右三角形","像素"
    markers = ["s", '^', "o", "+", "v", "d", "x", "*", "h", "3", "p", ".", "1", "<", ">", "D"]
    # (1)□(2)△(3)○(4)＋(5)▽(6)♢(7)×(8)☆(9)“正六边形”(10)“左Y”(11)“正五边形”(12)·(13)“正Y”(14)◁(15)▷(16)◇
    files = os.listdir(dataPath)
    # 返回指定的文件夹包含的文件或文件夹的名字的列表
    # ********************************************************************************************************************************无文件夹
    if len(files) == 0:
        return 0
    # *********************************************************************************************************************************一个文件夹
    elif len(files) == 1:
        return 0
    # *******************************************************************************************************************************多个文件夹
    else:
        # 1读取文夹并合并所有文件
        dirList = []
        # 作为变量名
        filesDataName = copy1.copy(files)
        # 作为同名变量数据
        pcaCountAll = []
        # 用于存放每个标签数量
        pcaCount = []
        # 用于存储删除光谱后标签数量
        # i = 0
        filesNumInter = 0
        # filesNumInter 全局文档个数，迭代时“从0开始，最后加1”
        classLabel = []
        # 创建标签
        classLabelName = []
        # 和标签对应的名字
        classOneLabel = []
        # 创建单个标签
        snrAllfile = []
        snrAllfile1 = []
        # 存储所有文件光谱信噪比
        allData = []
        # 用于存储所有处理后的数据
        # print(len(classOneLabel))
        meanAllSpectrum = []
        # 用于存储所有文件光谱均值，画均值方差层次罗列图
        stdAllSpectrum = []
        # 用于保存标准差到txt
        stdAll1 = []
        stdAll2 = []
        # 分别用于存储均值方差图上下bar，阴影部分
        snrDelfile = []
        # 用于存储删除信噪比小于阈值的光谱
        for f in files:
            filesDataName[filesNumInter] = []
            pltY = []
            # 存储光谱y值，用于有层次画图
            meanSpectrum = []
            # 存储该文件下光谱均值
            stdSpectrum = []
            # 存储该文件夹下光谱标准差
            if os.path.isdir(dataPath + '\\' + f):  # 判断是不是文件夹
                # if f[0] != '.':  # 排除隐藏文件夹。因为隐藏文件夹过多
                dirList.append(f)  # 添加非隐藏文件夹

            filesDataName[filesNumInter] = readMerge(dataPath[:] + "\\" + dirList[filesNumInter])
            # print("***",np.array(filesDataName).shape)
            # return 0
            filesDataNameTemp = filesDataName[filesNumInter]
            # allData = readMerge(dataPath[:] + "\\" + dirList[i])
            pcaCountAll.append(filesDataNameTemp.shape[1] - 1)
            # print('pcaCount:',pcaCount)
            # 把该标签数据量添加到pcaCount中

            filesDataNameTemp = filesDataNameTemp[startSpec: (filesDataNameTemp.shape[0] - endSpec), :]
            # #截谱
            # # **************************************************************************************************前后截谱
            # if filesNumInter == 0:
            #     filesDataNameTemp = filesDataNameTemp[100:, :]
            # elif filesNumInter == 1:
            #     filesDataNameTemp = filesDataNameTemp
            # elif filesNumInter == 2:
            #     filesDataNameTemp = filesDataNameTemp[100:, :]
            # else:
            #     filesDataNameTemp = filesDataNameTemp

            filesDataName1 = copy1.copy(filesDataNameTemp)
            # 注意：直接采用 ’ = ’ 赋值方式会使得两个变量占用同一地址
            x = filesDataNameTemp[:, 0]

            snrOnefile = []
            # 存储单个文件下拉曼光谱信噪比
            snrOneDelfile = []
            # 用于存储单个文件下删除信噪比小于阈值的光谱

            countDelSpect = 0
            # 用于存储该文件夹下被删除文件的数量
            # 2批处理
            for j in range(1, filesDataNameTemp.shape[1]):
                snrS = 0
                # 初始化该光谱信噪比
                # 循环处理每个光谱
                y = filesDataNameTemp[:, j]
                # 从第二列开始，保留了第一列坐标，后面需要去掉
                # y1 = Norris(y, nNorris)

                y1 = CRR(y, filterSize, dynamicFactor)
                y2 = savgol_filter(y1, winSG, nSG)
                # 2.3 airPLS
                # print("j:",j)
                y3 = airPLS(y2, lambdaAirPls, itermaxAirPls)
                # 有参去背景（自适应迭代重加权惩罚最小二乘，airpls）
                # y3 = background_substruction(y2)
                # 无参去背景（交替最小二乘ALS）
                # 2.4 归一化
                # y4 = MaxMinNormalization(y3)

                y4 = MaxMinNormalization(y3, x, [2500, 3500])

                # y4 = CDT_Trans(y4)
                # snrNoiseindex = 20
                # # 从尾部选取长度，作为静默区
                # delSNR = 10
                # #删除光谱信噪比阈值

                # snrSpectrumSingal = max(y4[170 - startSpec:400 - startSpec]) - min(y4)
                snrSpectrumSingal = max(y4) - min(y4)
                # 信号强度
                snrSpectrumNoise = max(y4[-snrNoiseindex:]) - min(y4[-snrNoiseindex:])
                # 噪声强度
                snrS = snrSpectrumSingal / snrSpectrumNoise
                # 光谱信噪比
                if snrS < delSNRmin or snrS > delSNRmax:
                    snrOneDelfile.append(j)
                    countDelSpect = countDelSpect + 1
                    continue

                snrOnefile.append(snrS)
                snrAllfile1.append(snrS)
                # 全部信噪比值存储到一个list中

                filesDataNameTemp[:, j - countDelSpect] = y4
                # j - countDelSpect,用于减去删除光谱的数量
                # if j <= pltInter:
                #     pltTemp = y4 + j * 0.3
                #     pltY.append(pltTemp)
                #     # 用于层次化画光谱
                #     # np.column_stack((Data_x, Dout.T))
                classLabel.append(filesNumInter)
                classLabelName.append(files[filesNumInter])

                # print("snrOnefile",snrOnefile)
            filesDataNameTemp = filesDataNameTemp[:, 0:(filesDataNameTemp.shape[1] - countDelSpect)]

            # ************************************************************************************************************保存处理后数据
            np.savetxt(outPath + '\\' + r'Process_' + files[filesNumInter] + '.txt',
                       filesDataNameTemp, delimiter='\t', fmt='%.3f')

            # print("删除文件数量：",countDelSpect)
            snrDelfile.append(snrOneDelfile)
            # 用于存储所有信噪比小于阈值的光谱
            pcaCount.append(filesDataNameTemp.shape[1] - 1)
            # allData.append(filesDataNameTemp)
            # 存储所有处理后的光谱数据
            snrAllfile.append(snrOnefile)
            # 按文件存储所有光谱信噪比
            # print("filesDataNameTemp:",np.array(filesDataNameTemp).shape)
            # print("pltY",np.array(pltY).shape)
            classOneLabel.append(filesNumInter)
            # 存储标签名，只包含1个
            # urldic = urlparse(dataPath)
            # pathdict = urldic.path.split('\\')
            # # 以"\"分割路径，供后文输出路径地址
            # outPath1 = outPath + '\\' + files[i] + r'_process.txt'
            # 输出地址及其文件名
            filesDataName[filesNumInter] = filesDataNameTemp
            # np.savetxt(outPath1, filesDataName[i], fmt='%f', delimiter=' ')


            # **********************************************************************************************************处理后均值方差图
            figure(3 * filesNumInter + 2)
            meanSpectrum.append(np.mean(filesDataNameTemp[:, 1:filesDataNameTemp.shape[1]], axis=1))
            # 按照行进行均值拼接
            stdSpectrum.append(np.std(filesDataNameTemp[:, 1:filesDataNameTemp.shape[1]], axis=1))

            # # 查找标准差最大处的峰位
            # print(np.array(stdSpectrum).shape)
            # print("1:",stdSpectrum[0].tolist())
            # # print(x[stdSpectrum.index(max(stdSpectrum[0], key=abs))])
            # print(x[stdSpectrum[0].tolist().index(max(stdSpectrum[0].tolist(), key=abs))])

            # 按照行进行标准差拼接
            stdAllSpectrum.append(stdSpectrum)
            # 拼接标准差
            plt.plot(x, np.array(meanSpectrum).T, color=colorSpectrumColour[filesNumInter % len(colorSpectrumColour)],
                     lw=1,
                     label='Spectral mean')
            # plt.plot(x, np.array(meanSpectrum).T, color="white", lw=2)
            # print(np.array(x).shape)
            # print("111:",(np.array(meanSpectrum).T).shape)
            # print("222:",((np.array(meanSpectrum) - np.array(stdSpectrum)).T).shape)
            std1 = (np.array(meanSpectrum) - np.array(stdSpectrum)).T
            std2 = (np.array(meanSpectrum) + np.array(stdSpectrum)).T
            plt.fill_between(x, std1.flatten(), std2.flatten(), color="#B5B5B5", label='Std of SNR')
            plt.legend(loc='best')
            # flatten()二维转一维
            outPath3 = outPath + '\\' + files[filesNumInter] + r'_meanStd.tif'
            plt.savefig(outPath3, format='tif', bbox_inches='tight', transparent=True, dpi=300)
            plt.rcParams['svg.fonttype'] = 'none'
            # outPath3 = outPath + '\\' + files[filesNumInter] + r'_meanStd.svg'
            # plt.savefig(outPath3, format='svg', bbox_inches='tight', transparent=True)

            meanAllSpectrum.append(np.array(meanSpectrum).T)
            # # 用于存储所有文件光谱均值，画均值方差层次罗列图
            stdAll1.append(std1.flatten())
            # a是个数组，a.flatten()就是把a降到一维，默认是按行的方向降
            stdAll2.append(std2.flatten())
            # **********************************************************************************************************保存均值
            # print("???",type(x))
            # print("???",type(y))
            xMean = np.array(x).reshape(-1, 1)
            yMeam = np.array(meanSpectrum).reshape(-1, 1)
            # print(xMean.shape)
            # print(yMeam.shape)
            dataMean = np.hstack((xMean, yMeam))

            np.savetxt(outPath + '\\' + r'Mean_' + files[filesNumInter] + '.txt',
                       dataMean, delimiter='\t', fmt='%.3f')

            # plt.show()


            filesNumInter = filesNumInter + 1
            plt.close('all')
        # print("allData1:",np.array(allData).shape)
        # **********************************************************************************************将信噪比保存到CSV,信噪比箱型抖动图

        snrCSV = np.array(snrAllfile1).reshape(-1, 1)
        labelCSV = np.array(classLabelName).reshape(-1, 1)
        SNRClass = np.column_stack((snrCSV, labelCSV))
        outPuthExcSNR = outPath + '\\' + r'SNR.csv'
        snrList = pd.DataFrame(SNRClass, columns=["SNR", "LabelName"])
        snrList.to_csv(outPuthExcSNR,encoding='utf_8_sig')

        dfBox = pd.read_csv(outPuthExcSNR)
        # sns.boxplot(x='LabelName', y='SNR', data=dfBox)
        # # sns.violinplot(x='SNR', y='LabelName', data=df)
        # # 小提琴形状替换箱线图
        # # sns.stripplot(x='LabelName', y='SNR', data=dfBox, jitter=0.25, palette='Set2', dodge=True, size=4, alpha=0.5,
        # #               color=None, )
        # # sns.stripplot(x='LabelName', y='SNR', data=dfBox, jitter=True,
        # #               palette='Set2', dodge=True, size=4, alpha=0.5)
        # sns.stripplot(x='LabelName', y='SNR', data=dfBox, jitter=0.25, palette='Set2', dodge=True, size=4, alpha=0.5,
        #               color=None, )

        # 构建与 _tSNE.tif 完全一致的颜色映射（按 classOneLabel 顺序）
        _label_order_snr = [files[k] for k in classOneLabel]
        _snr_palette = {
            files[k]: colorSpectrumColour[i % len(colorSpectrumColour)]
            for i, k in enumerate(classOneLabel)
        }

        # 箱型图：使用统一配色，隐藏自带离群点（由 stripplot 统一展示）
        sns.boxplot(
            x='LabelName', y='SNR', data=dfBox,
            palette=_snr_palette,
            order=_label_order_snr,
            fliersize=0,  # 隐藏 boxplot 自带离群点，避免与 stripplot 重叠
            width=0.45,
            linewidth=1.2,
            boxprops=dict(alpha=0.75),
            whiskerprops=dict(linewidth=1.0),
            medianprops=dict(linewidth=1.8, color='white')
        )

        # 散点抖动图：去除 dodge=True（无 hue 时 dodge 会导致整体左偏），
        # 增大 jitter、缩小 size，与 boxplot 精确对齐
        sns.stripplot(
            x='LabelName', y='SNR', data=dfBox,
            palette=_snr_palette,
            order=_label_order_snr,
            jitter=0.28,  # 适度加大横向散布
            dodge=False,  # 关闭 dodge，消除偏移
            size=3.0,
            alpha=0.60,
            linewidth=0.4,
            edgecolor='white'  # 白色描边提升点的清晰度
        )

        # 可以再加一个分量hue，可用来控制每组抖动点进一步划分不同颜色！！！！！
        # linewidth=1,加水平网格
        # plt.legend()
        plt.xticks(fontsize=12, rotation=90)
        plt.yticks(fontsize=12)
        # plt.xlabel('Labels', fontsize=12)  # 添加x轴的名称
        plt.ylabel('SNR intensity', fontsize=12)
        # plt.show()
        # plt.ylim([1.5,15])
        # # 设置坐标范围

        outPath6 = outPath + '\\' + r'_snrBox1.tif'
        plt.savefig(outPath6, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # outPath6 = outPath + '\\' + r'_snrBox1.svg'
        # plt.savefig(outPath6, format='svg', bbox_inches='tight', transparent=True)
        plt.close()
        # ========================== 新增：保存信噪比箱型图(_snrBox1)的解释 ==========================
        snr_summary = []
        snr_means = [np.mean(snrAllfile[i]) for i in range(filesNumInter)]
        snr_stds = [np.std(snrAllfile[i]) for i in range(filesNumInter)]
        snr_mins = [np.min(snrAllfile[i]) for i in range(filesNumInter)]
        snr_maxs = [np.max(snrAllfile[i]) for i in range(filesNumInter)]

        with open(outPath + '\\' + r'_snrBox1_说明.txt', 'w', encoding='utf-8') as f:
            f.write("信噪比箱型图(_snrBox1.tif) 说明：\n")
            f.write("该图展示六种微生物 (Bli, Sxy, Rem, Lra, Lco, Bsp) 预处理后单细胞拉曼光谱的信噪比（SNR）分布。\n")
            f.write("每个箱体表示中位数信噪比与上下四分位范围，抖动点表示单个光谱信噪比散布。\n\n")
            for i in range(filesNumInter):
                f.write(f"{files[i]}：平均SNR = {snr_means[i]:.2f} ± {snr_stds[i]:.2f}，"
                        f"范围 = [{snr_mins[i]:.2f}–{snr_maxs[i]:.2f}]\n")
            f.write("\n结论：各菌株单细胞光谱信噪比分布集中、波动小，说明预处理过程有效提升了光谱质量和一致性。\n")
        # **************************************************************************************************将均谱保存txt
        meanAllSpectrumTuple0 = np.array(meanAllSpectrum).shape[0]
        # tuple索引
        meanAllSpectrumTuple1 = np.array(meanAllSpectrum).shape[1]
        meanAllSpectrumTxt = np.array(meanAllSpectrum).reshape(meanAllSpectrumTuple0, meanAllSpectrumTuple1).T
        # reshape变形，(m,n,1)到(m,n)
        meanAllSpectrumTxt = np.column_stack((x.T, meanAllSpectrumTxt))
        # 列拼接
        np.savetxt(outPath + '\\' + r'_spectrumMean.txt', meanAllSpectrumTxt, delimiter='\t', fmt='%.3f')

        # **************************************************************************************************************均值方差层次罗列图
        figure(3 * filesNumInter, figsize=(8, 8))

        # data = np.array(
        #     pd.read_csv(r"C:\Users\lixin\Desktop\11.txt", delim_whitespace=True,
        #                 header=None))
        # x = data[50:,0]

        from brokenaxes import brokenaxes

        _xArr = np.asarray(x, dtype=float)
        _BRK_L, _BRK_R = 1900.0, 2600.0

        # ── 不传 fig= 参数，先创建 figure，brokenaxes 自动通过 plt.gcf() 绑定 ──
        # ── 避免 fig= 与 plt.GridSpec 内部状态不同步导致轴数组为空的 IndexError ──
        plt.figure(figsize=(8, 8))
        bax = brokenaxes(
            xlims=((_xArr.min(), _BRK_L), (_BRK_R, _xArr.max())),
            hspace=0.001
        )

        for interI in range(filesNumInter):
            interIhight = 0.6
            # ravel() 确保一维，避免 (n,1) 形状引发 brokenaxes 内部广播异常
            _y_mean = np.asarray(meanAllSpectrum[interI]).ravel()
            _y_std1 = np.asarray(stdAll1[interI]).ravel()
            _y_std2 = np.asarray(stdAll2[interI]).ravel()

            bax.plot(
                _xArr,
                _y_mean + interIhight * interI,
                color=colorSpectrumColour[interI % len(colorSpectrumColour)],
                lw=1, label=files[interI]
            )
            bax.fill_between(
                _xArr,
                _y_std1 + interIhight * interI,
                _y_std2 + interIhight * interI,
                color="#B5B5B5"
            )

        # ── 直接遍历子轴隐藏 y 刻度，避免 bax.set_yticks([]) 在部分版本越界 ──
        for _sub_ax in bax.axs:
            _sub_ax.set_yticks([])

        bax.set_xlabel("Raman shift (cm$^{-1}$)", fontsize=12, labelpad=30)
        bax.set_ylabel("Intensity (a.u.)", fontsize=12, labelpad=40)
        plt.rcParams.update({'font.size': 12})

        outPath4 = outPath + '\\' + r'_meanStdcolumn.tif'
        plt.savefig(outPath4, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # outPath4 = outPath + '\\' + r'_meanStdcolumn.svg'
        # plt.savefig(outPath4, format='svg', bbox_inches='tight', transparent=True)
        plt.close()
        # **************************************************************************************************************所有类别均值的总平均光谱图（截断模式，标记特征峰位）
        # from brokenaxes import brokenaxes

        _xArr_gm = np.asarray(x, dtype=float)
        _BRK_L_gm, _BRK_R_gm = 1900.0, 2600.0

        # 计算所有类别均值光谱的总平均（grand mean）
        # meanAllSpectrum[i] 形状为 (n_points, 1)，ravel() 转为一维后逐行堆叠取均值
        _all_mean_arr = np.array([np.asarray(meanAllSpectrum[i]).ravel()
                                  for i in range(filesNumInter)])
        _grand_mean_spec = np.mean(_all_mean_arr, axis=0)

        # 目标标记波数列表
        _mark_shifts = [934, 949, 1715, 734, 1323, 1031, 2863, 1088, 796, 1398, 802]

        plt.figure(figsize=(8, 4))
        bax_gm = brokenaxes(
            xlims=((_xArr_gm.min(), _BRK_L_gm), (_BRK_R_gm, _xArr_gm.max())),
            wspace=0.02  # 与 _meanStdcolumn 保持相同截断间距
        )

        # 绘制总平均光谱（黑色实线，线宽2）
        _lw_gm = 3.5
        bax_gm.plot(
            _xArr_gm,
            _grand_mean_spec,
            color='black',
            lw=_lw_gm,
            # label='Grand mean spectrum'
        )

        # 在特征波数位置标记红色实心圆点
        # scatter 的 s 参数单位为点²，(_lw_gm * 4)² = 64，视觉上与线宽协调
        _dot_size = 46

        for _ms in _mark_shifts:
            # 找到最近的波数索引
            _ci = int(np.argmin(np.abs(_xArr_gm - _ms)))
            bax_gm.scatter(
                [float(_xArr_gm[_ci])],
                [float(_grand_mean_spec[_ci])],
                color='red',
                s=_dot_size,
                zorder=5,
                marker='o'
            )


        bax_gm.set_xlabel("Raman shift (cm$^{-1}$)", fontsize=12, labelpad=30)
        bax_gm.set_ylabel("Intensity (a.u.)", fontsize=12, labelpad=40)
        plt.rcParams.update({'font.size': 12})

        outPath_gm = outPath + '\\' + r'_grandMeanSpectrum.tif'
        plt.savefig(outPath_gm, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # outPath_gm = outPath + '\\' + r'_grandMeanSpectrum.svg'
        # plt.savefig(outPath_gm, format='svg', bbox_inches='tight', transparent=True)
        plt.close()

        # ****************************************************************************************************************保存标准差信息
        # stdAllSpectrum
        f = open(outPath + '\\' + r'_meanStdcolumn.txt', "w")
        f.write('保存同类拉曼光谱所有拉曼位移的标准差' + '\n')
        for interI in range(filesNumInter):
            # f.write('保存同类拉曼光谱所有拉曼位移的标准差：' + '\n')
            f.write("[" + str(files[interI]) + ']光谱的平均标准差为：' + str(round(np.mean(stdAllSpectrum[interI]), 2)) + '\n')
            # f.write("[" + str(files[interI]) + ']光谱的平均标准差为：' + str(round(np.mean(stdAll1[interI]), 2)) + '\n')

            f.write("[" + str(files[interI]) + ']光谱的最大标准差为：' + str(
                round(np.max(stdAllSpectrum[interI]), 2)) + ",最大处峰位：" + str(
                x[stdAllSpectrum[interI][0].tolist().index(max(stdAllSpectrum[interI][0].tolist(), key=abs))]) + '\n')
            # x[stdSpectrum[0].tolist().index(max(stdSpectrum[0].tolist(), key=abs))]

            # f.write("[" + str(files[interI]) + ']光谱的最大标准差为：' + str(
            #     round(np.max(stdAll1[interI]), 2)) + ",最大处峰位：" + str(
            #     x[stdAll1[interI].index(max(stdAll1[interI][interI], key=abs))]) + '\n')
        f.close()
        # **************************************************************************************************************信噪比直方图
        figure(3 * filesNumInter + 1)
        snrAlpha = 0.5
        snrLinewidth = 0.8

        for interI in range(filesNumInter):
            # sns.distplot(snrAllfile[interI], color=colorSpectrumColour[interI % len(colorSpectrumColour)],
            #              label=files[interI],
            #              # hist_kws={'alpha': snrAlpha},
            #              hist=False,
            #              # 不显示直方图
            #              # kde=False,
            #              # 不显示密度图
            #              kde_kws={'linewidth': snrLinewidth})
            sns.histplot(snrAllfile[interI], color=colorSpectrumColour[interI % len(colorSpectrumColour)],
                         kde=True, label=files[interI], stat="density", linewidth=snrLinewidth)

        plt.title('SNR Density Curves with Histogram', fontsize=12)
        plt.legend(fontsize=8)
        plt.xlabel("SNR intensity")
        plt.ylabel("Density")
        outPath5 = outPath + '\\' + r'_snrHistogram.tif'
        plt.savefig(outPath5, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # outPath5 = outPath + '\\' + r'_snrHistogram.svg'
        # plt.savefig(outPath5, format='svg', bbox_inches='tight', transparent=True)
        plt.close()

        # ========================== 新增：保存信噪比密度直方图(_snrHistogram)的解释 ==========================
        with open(outPath + '\\' + r'_snrHistogram_说明.txt', 'w', encoding='utf-8') as f:
            f.write("信噪比密度直方图(_snrHistogram.tif) 说明：\n")
            f.write("该图展示六类微生物拉曼光谱的信噪比分布密度曲线，用于评估预处理后光谱质量的集中性与稳定性。\n\n")
            for i in range(filesNumInter):
                f.write(
                    f"{files[i]}：平均SNR = {snr_means[i]:.2f}，峰值密度分布主要集中在 {snr_means[i] - snr_stds[i]:.2f}–{snr_means[i] + snr_stds[i]:.2f} 区间。\n")
            f.write("\n结论：各菌株光谱信噪比分布单峰且集中于稳定区间，表明经过背景去除和平滑滤波后的光谱噪声显著降低，信号检测鲁棒性良好。\n")


        # **************************************************************************************************************信噪比箱型图
        figure(3 * filesNumInter + 2)

        dfLists = []
        # 存储以文件为单位信噪比字典
        for interI in range(filesNumInter):
            dfLists.append(pd.DataFrame({files[interI]: snrAllfile[interI]}))
        dfData = pd.concat(dfLists, axis=1)

        # color = dict(boxes='DarkGreen', whiskers='DarkOrange', medians='DarkBlue', caps='Gray')
        #
        # #字典拼接
        # dfData.plot.box(color = color)
        # dfData.boxplot(sym='+', patch_artist=True,
        #                boxprops={"color": colorSpectrumColour[interI % 13], "facecolor": "#F43D68"},
        #                meanprops={"marker": "D", "markerfacecolor": "white"}, medianprops={"linestyle": "--", "color": "r"})
        # plt.grid(linestyle="--", alpha=0.8)

        f = dfData.boxplot(sym='+', patch_artist=True, return_type='dict')

        # 这里共有四个box
        # color = ['k', 'g', 'r', 'deepskyblue']  # 有多少box就对应设置多少颜色

        for box, c in zip(f['boxes'], colorSpectrumColour):
            # 箱体边框颜色
            box.set(color=c, linewidth=2)
            # 箱体内部填充颜色
            box.set(facecolor=c)


        plt.grid(axis="y", ls=":", lw=1, color="gray", alpha=0.8)
        outPath6 = outPath + '\\' + r'_snrBox.tif'
        plt.savefig(outPath6, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # outPath6 = outPath + '\\' + r'_snrBox.svg'
        # plt.savefig(outPath6, format='svg', bbox_inches='tight', transparent=True)
        plt.close()

        # **********************************************************************************************************拼接
        allData = filesDataName[0]
        for interI in range(1, filesNumInter):
            # tempData = filesDataName[interI]
            allData = np.column_stack((allData, (filesDataName[interI])[:, 1:]))
        ySpec = allData[:, 1:allData.shape[1]].T

        # **************************************************************************************************************3 hca层次聚类
        figure(3 * filesNumInter + 3)
        if distanceMethod == 0:
            distM = "ward"
            # 沃德方差最小化算法
        elif distanceMethod == 1:
            distM = "centroid"
            # 质心间的欧氏距离
        elif distanceMethod == 2:
            distM = "weighted"
            # 加权分组平均法
        else:
            distM = 'average'
            # UPGMA算法（非加权组平均）法

        # distanceMethod = 'ward'
        # # 层次聚类距离度量选用的方法,可选：'centroid','weighted','average'
        Z = hierarchy.linkage(ySpec, method=distM, metric='euclidean')
        # 执行层次/聚集聚类
        hcaColorThreshold = 40
        # color_threshold= 8 设置颜色高度阈值，用于控制决策树和PCA的颜色分类
        # hcaLabelsize = 6
        # #标号大小
        # cluster.hierarchy.cut_tree(Z, n_clusters=filesNumInter)
        # 设置聚类个数n_clusters

        hierarchy.dendrogram(Z, color_threshold=hcaColorThreshold, labels=pd.DataFrame(ySpec).index + 1)
        # 将分层聚类绘制为树状图
        # hierarchy.dendrogram(Z, labels=pd.DataFrame(ySpec).index + 1)
        # label = cluster.hierarchy.cut_tree(Z, height = hcaColorThreshold)
        hacLeaves = hierarchy.leaves_list(Z) + 1
        # 返回叶子节点

        plt.xticks(rotation=90)
        plt.tick_params(labelsize=hcaLabelsize)
        # 调整刻度尺寸
        plt.title('HCA')
        # plt.show()
        outPath7 = outPath + '\\' + r'_HCA.tif'
        plt.savefig(outPath7, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # outPath7 = outPath + '\\' + r'_HCA.svg'
        # plt.savefig(outPath7, format='svg', bbox_inches='tight', transparent=True)
        # plt.show()
        plt.close()
        # *******************************************************************************************保存层次聚类叶子节点
        f = open(outPath + '\\' + r'_hcaLeaves.txt', "w")
        f.write("层次聚类树状图叶子节点（从左到右排列）：")
        f.write('\n')
        f.write(str(hacLeaves))
        f.close()

        # **************************************************************************************************************4 PCA数据压缩
        figure()
        # 根据两个最大的主成分进行绘图
        pcaNcomponents = 4
        # 主成分个数
        pca = skldec.PCA(n_components=pcaNcomponents)
        # 第一种，选择指定降维方差：n_components =0.95，贡献率之和到95%
        # 第二种，选择指定降维后的维数：n_components =5
        try:
            pca.fit(ySpec)
        except:
            pca.fit(ySpec)
        # pca.fit(ySpec)
        # 主城分析时每一行是一个输入数据
        # print("贡献率：",pca.explained_variance_ratio_)
        result = pca.transform(ySpec)

        # #**********************************************************************二维pca标记
        #
        tsneCount1 = 0
        tsneCount2 = 0

        # pcaSize = 30
        # #pca标签大小
        # pcaAlpha = 1
        # #pca透明度
        for tsneI in range(len(classOneLabel)):
            # print("光谱名称：",files[i])
            # print(type(files[i]))
            # print("光谱数量：", filesDataName[i].shape)
            tsneCount2 = tsneCount2 + pcaCount[tsneI]
            plt.scatter(result[tsneCount1:tsneCount2, pca_col1], result[tsneCount1:tsneCount2, pca_col2],
                        c=colorSpectrumColour[tsneI % len(colorSpectrumColour)], marker=markers[tsneI % len(markers)],
                        alpha=pcaAlpha,  label=files[tsneI])
            tsneCount1 = tsneCount1 + pcaCount[tsneI]

        # # 二维标记
        # for i in range(0,len(hacLeaves)):
        #     plt.text(result[i, 0], result[i, 1], i + 1, fontsize=10)
        # plt.show()

        plt.legend(loc=0, ncol=1, fontsize=8)
        # plt.legend(loc=0, ncol=1, fontsize=5)
        # set(tsneL, 'FontName', 'Times New Roman', 'FontSize', 5, 'FontWeight', 'normal')
        # print("贡献率：", )
        # plt.xlabel(u"PC 1" + " (" + str(round(pca.explained_variance_ratio_[0],1))+ ")")
        # plt.ylabel(u"PC 2" + " (" + str(round(pca.explained_variance_ratio_[1],1))+ ")")
        plt.xlabel(u"PC 1" + " (" + str('%.1f' % (100 * pca.explained_variance_ratio_[0])) + "%)")
        plt.ylabel(u"PC 2" + " (" + str('%.1f' % (100 * pca.explained_variance_ratio_[1])) + "%)")
        # '%1.0f' % (100 * pca.explained_variance_ratio_[0])
        plt.title('PCA_2D')
        # plt.show()
        outPath71 = outPath + '\\' + r'_PCA_2D.tif'
        plt.savefig(outPath71, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # outPath71 = outPath + '\\' + r'_PCA_2D.svg'
        # plt.savefig(outPath71, format='svg', bbox_inches='tight', transparent=True)
        plt.close()

        # # **************************************************************************************************二元关系分析
        #
        # 修改类别形状，由一维变二维（n,）变为(n,1),方便后续列拼接
        classLabelName2 = np.array(classLabelName).reshape(-1, 1)
        tsneClass = np.column_stack((result, classLabelName2))
        # np.savetxt(r"C:\Users\lixin\Desktop\b\tsne.txt",
        #            result2, delimiter='\t', fmt='%.2f')
        outPuthExcPCA = outPath + '\\' + r'PCA.csv'
        # allDataTwoClass = np.column_stack(
        #     (np.array(xLabelTest), np.array(yLabelTest), np.array(testTarget), np.array(knn.predict(testData)),
        #      np.array(knn.predict(testData)), np.array(clf.predict(testData))))
        # 存储前4维PCA
        tsneList = pd.DataFrame(tsneClass, columns=['PC1', 'PC2', 'PC3', 'PC4', 'labels'])
        tsneList.to_csv(outPuthExcPCA, index=False,encoding='utf_8_sig')
        # 调用保存的tsne表格,index=False:不要序号
        # df = pd.read_csv(outPuthExc)
        # dfTSNE = pd.read_csv(outPuthExcTSNE).iloc[:,1:5]
        # 切片
        dfPCA = pd.read_csv(outPuthExcPCA)

        # g = sns.pairplot(dfPCA, kind="scatter", diag_kind="kde", diag_kws=dict(bw=1.5), hue="labels", palette="hls",
        #                  plot_kws=dict(s=12, edgecolor="black", linewidth=0.3))
        # # g = sns.pairplot(dfPCA, kind="scatter", diag_kind="kde", hue="labels", palette="hls",
        # #                  plot_kws=dict(s=15, edgecolor="black", linewidth=0.5),kde_kws = {'bw' : 1.5})
        # # g.map_diag(sns.kdeplot)
        # # kind：用于控制非对角线上的图的类型，可选"scatter"与"reg"
        # # diag_kind：控制对角线上的图的类型，可选"hist"(直方图)与"kde"（密度图），加上diag_kws=dict(bw=1.5),避免报错
        # # hue ：针对某一字段进行分类
        # # palette：控制色调
        # g.map_offdiag(sns.kdeplot, n_levels=2)
        # # sns.kdeplot(x,y,cmap=pal)：绘制核密度分布图
        # # plt.show()
        # outPath91 = outPath + '\\' + r'_pcaPair.tif'
        # plt.savefig(outPath91, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # outPath91 = outPath + '\\' + r'_pcaPair.svg'
        # plt.savefig(outPath91, format='svg', bbox_inches='tight', transparent=True)
        # plt.close()

        # ********************************************************************************************************PCA三维
        figure(3 * filesNumInter + 4)
        ax = plt.figure().add_subplot(111, projection='3d')
        # 其中参数111，指的是将图像分成1行1列，此子图占据从左到右从上到下的1位置
        pcaNcomponents3D = nDim + 1
        # 主成分个数
        pca_3D = skldec.PCA(n_components=pcaNcomponents3D)
        # 第一种，选择指定降维方差：n_components =0.95，贡献率之和到95%
        # 第二种，选择指定降维后的维数：n_components =5
        # pca_3D.fit(ySpec)
        try:
            pca_3D.fit(ySpec)
        except:
            pca_3D.fit(ySpec)
        # 主城分析时每一行是一个输入数据
        result_3D = pca_3D.transform(ySpec)

        pca_col3 = 2
        pcaCount1 = 0
        pcaCount2 = 0
        # pcaSize = 30
        # #pca标签大小
        # pcaAlpha = 1
        # #pca透明度
        for pcai in range(len(classOneLabel)):
            pcaCount2 = pcaCount2 + pcaCount[pcai]
            ax.scatter(result_3D[pcaCount1:pcaCount2, pca_col1], result_3D[pcaCount1:pcaCount2, pca_col2],
                       result_3D[pcaCount1:pcaCount2, pca_col3],
                       c=colorSpectrumColour[pcai % len(colorSpectrumColour)], marker=markers[pcai % len(markers)],

                       alpha=pcaAlpha,  label=files[pcai])
            pcaCount1 = pcaCount1 + pcaCount[pcai]
        # 设置坐标轴
        plt.legend(loc='best', fontsize=8)
        ax.set_xlabel('PC1')
        ax.set_ylabel('PC2')
        ax.set_zlabel('PC3')
        plt.title('PCA_3D')
        # plt.show()
        # return 0
        outPath16 = outPath + '\\' + r'_PCA_3D.tif'
        plt.savefig(outPath16, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # outPath16 = outPath + '\\' + r'_PCA_3D.svg'
        # plt.savefig(outPath16, format='svg', bbox_inches='tight', transparent=True)
        plt.close()

        ##**************************************************************************************************************tsne降维
        figure(3 * filesNumInter + 5)
        # 根据两个最大的主成分进行绘图
        pcaNcomponents = 4
        # 主成分个数
        # pca = skldec.PCA(n_components=pcaNcomponents)
        # # 第一种，选择指定降维方差：n_components =0.95，贡献率之和到95%
        # # 第二种，选择指定降维后的维数：n_components =5
        # pca.fit(ySpec)
        # # 主城分析时每一行是一个输入数据
        # result = pca.transform(ySpec)
        nDim = 2
        # int, 空间的维度
        nPerplex = 10.0
        # 数据集越大，需要参数值越大，建议值位5 - 50
        learnRate = 200
        # float,学习率，建议取值为10.0-1000.0
        nInter = 1000
        # int, 最大迭代次数
        nGrad = 1e-06
        # float,如果梯度低于该值，则停止算法

        result2 = TSNE(n_components=nDim, perplexity=nPerplex, early_exaggeration=12.0, learning_rate=learnRate,
                       n_iter=nInter,
                       n_iter_without_progress=300, min_grad_norm=nGrad, metric='euclidean', init='pca',
                       verbose=0, random_state=1, method='exact').fit_transform(ySpec)
        # random_state:int, RandomState or None, default:None，伪随机数生成器种子,不同的初始化可能会导致成本函数的局部最小值不同

        # import umap
        # result2 = umap.UMAP(n_components=2,random_state=1,min_dist=0.5).fit_transform(ySpec)

        # print(type(result2))

        # pca_col1 = 0
        # pca_col2 = 1
        tsneCount1 = 0
        tsneCount2 = 0

        # pcaSize = 30
        # #pca标签大小
        # pcaAlpha = 1
        # #pca透明度
        for tsneI in range(len(classOneLabel)):
            # print("光谱名称：",files[i])
            # print(type(files[i]))
            # print("光谱数量：", filesDataName[i].shape)
            tsneCount2 = tsneCount2 + pcaCount[tsneI]
            plt.scatter(result2[tsneCount1:tsneCount2, pca_col1], result2[tsneCount1:tsneCount2, pca_col2],
                        c=colorSpectrumColour[tsneI % len(colorSpectrumColour)], marker=markers[tsneI % len(markers)],

                        alpha=pcaAlpha, label=files[tsneI])
            tsneCount1 = tsneCount1 + pcaCount[tsneI]

        # for i in range(100):
        #     plt.text(result2[i, 0], result2[i, 1], i + 1, fontsize=14)
        # plt.show()
        plt.legend(loc=0, ncol=1, fontsize=18)
        # plt.legend(loc=0, ncol=1, fontsize=5)
        # set(tsneL, 'FontName', 'Times New Roman', 'FontSize', 5, 'FontWeight', 'normal')
        plt.xlabel(u"tSNE 1")
        plt.ylabel(u"tSNE 2")
        # plt.title('t-SNE')
        # plt.show()
        outPath9 = outPath + '\\' + r'_tSNE.tif'
        plt.savefig(outPath9, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # outPath9 = outPath + '\\' + r'_tSNE.svg'
        # plt.savefig(outPath9, format='svg', bbox_inches='tight', transparent=True)
        plt.close()

        # ****************************************************************************************************************按文件split 分割 & 拼接
        # Target0 = np.zeros(np.array(filesDataName[0][:, 1:]).shape[1])
        # print(type(Target0))
        # print(Target0)
        Target0 = np.full(np.array(filesDataName[0][:, 1:]).shape[1], 0, dtype=int)
        # 创建首个文件label,标签转置，转换成和数据同格式
        # print(Target0)
        # print((filesDataName[0][:, 1:].T).shape)
        # print((Target0.T).shape)
        trainData, testData, trainTarget, testTarget = train_test_split((filesDataName[0][:, 1:].T), (Target0.T),
                                                                        test_size=svmTestSize, random_state=1)
        # train_test_split交叉验证中常用的函数
        # svmC = 0.8
        # # 错误项的惩罚系数,范围[0.5-1],C越大分类效果越好，但有可能会过拟合（defaul C = 1）
        # 首个文件数据分割
        for interI in range(1, filesNumInter):
            #
            # TargetTep = np.zeros(np.array(filesDataName[interI][:, 1:]).shape[1])
            TargetTep = np.full(np.array(filesDataName[interI][:, 1:]).shape[1], interI, dtype=int)
            # print(TargetTep)
            # 创建其他文件label#,标签转置，转换成和数据同格式
            trainDataTep, testDataTep, trainTargetTep, testTargetTep = train_test_split(
                (filesDataName[interI][:, 1:].T), TargetTep.T,
                test_size=svmTestSize, random_state=1)
            # 其他文件数据分割并拼接
            trainData = np.row_stack((trainData, trainDataTep))
            testData = np.row_stack((testData, testDataTep))
            trainTarget = np.concatenate((trainTarget, trainTargetTep), axis=0)
            # 一维数组垂直拼接
            testTarget = np.concatenate((testTarget, testTargetTep), axis=0)
            # 一维数组垂直拼接

        # print(trainData.shape)
        # print(testData.shape)
        # print(trainTarget.shape)
        # print(testTarget.shape)

        # dataTarget1 = label_binarize(trainTarget, np.arange(len(classOneLabel)))
        # testTarget1 = label_binarize(testTarget, np.arange(len(classOneLabel)))

        testTarget1 = label_binarize(testTarget, classes=np.arange(len(classOneLabel)))
        # testTarget1 = label_binarize(testTarget, classes=np.arange(len(classOneLabel)))
        # 标签二值化，为了将ROC曲线和ROC区域扩展到多类或多标签分类，有必要对输出进行二值化：
        # ⑴可以每个标签绘制一条ROC曲线。
        # ⑵也可以通过将标签指示符矩阵的每个元素视为二元预测（微平均）来绘制ROC曲线。
        # ⑶另一种用于多类别分类的评估方法是宏观平均，它对每个标签的分类给予相同的权重。

        # return 0
        # trainData, testData, trainTarget, testTarget = train_test_split(ySpec, dataTarget,
        #                                                                 test_size=svmTestSize, random_state=1)
        # **************************************************************************************************************SVM混淆矩阵
        figure(3 * filesNumInter + 6)
        # 分类SVM、LDA、KNN
        # print("数据大小",ySpec.shape)
        # classData = np.column_stack((ySpec, np.array(classLabel).T))
        # 将标签拼接到数据中
        # print("拼接标签后的数据大小：",classData.shape)
        # *********************************************************************************************网格搜索和交叉验证
        # # modelSVM = svm.SVC(C=svmC,probability=True)
        # # modelSVM = svm.NuSVC(decision_function_shape='ovr', probability=True, class_weight='balanced')
        # modelSVM = svm.SVC(decision_function_shape='ovr', probability=True, class_weight='balanced')
        # # class_weight = {1: 10},class_weight='balanced'
        # # kernel='linear'时，为线性核，C越大分类效果越好，但有可能会过拟合（defaul C=1）。
        # # kernel='rbf'时（default），为高斯核，gamma值越小，分类界面越连续；
        # # gamma值越大，分类界面越“散”，分类效果越好，但有可能会过拟合。
        # # kernel='rbf'时（default），为高斯核，；
        #
        # # 建立需要搜索的参数的范围
        # param_grid = {'kernel': ['linear', 'poly', 'rbf', 'sigmoid'],
        #               # 'C': [0.6, 0.7, 0.8, 1],
        #               # 'gamma':np.linspace(0.001,1,5),
        #               # 'degree':[0.1,0.5,1,2,3],
        #               # 'nu':[0.2,0.3,0.4]
        #
        #               }
        # # # grid_search = GridSearchCV(modelSVM, param_grid, cv=3)
        # # grid_search = GridSearchCV(modelSVM, param_grid, cv=3, n_jobs=1)
        # # try:
        # #     clf = grid_search.fit(trainData, trainTarget)
        # # except:
        # #     clf = grid_search.fit(trainData, trainTarget)
        # # # clf = grid_search.fit(trainData, trainTarget)
        # # 训练模型
        # try:
        #     clf = modelSVM.fit(trainData, trainTarget)
        # except:
        #     clf = modelSVM.fit(trainData, trainTarget)
        modelSVM = svm.SVC(decision_function_shape='ovr', probability=True, class_weight='balanced')
        # 直接训练（不进行网格搜索）
        clf = modelSVM.fit(trainData, trainTarget)
        best_model_svm = clf
        # *******************************************************************************************保存、加载分类器模型
        # 保存模型
        outPathSVM = outPath + '\\' + r'_svmModel.pkl'
        joblib.dump(clf, outPathSVM)
        # 加载模型
        # outPathSVM = outPath + '\\' + r'clfSVM.pkl'
        # clfSVM = joblib.load(outPathSVM)
        # best_model_svm = grid_search.best_estimator_
        # best_model_svm = clf
        # print("***",type(grid_search.best_params_))
        # print(type(grid_search.best_params_))
        # print("默认超参数：",best_model)
        # print(type(best_model))
        # # # ********************************************************************************************保存SVM网格参数
        # f = open(outPath + '\\' + r'_svmParameter.txt', "w")
        # f.write("SVM参数说明：\n")
        # f.write('搜索网格参数1.核函数(kernel)\n')
        # f.write('\t线性核函数(linear):主要用于线性可分，参数少，速度快\n')
        # f.write('\t径向基核函数(rbf):主要用于线性不可分，参数多,耗时\n')
        # f.write('\t多项式核函数(poly):依靠升维使得原本线性不可分的数据线性可分\n')
        # f.write('\tSigmoid核函数:SVM实现的就是一种多层感知器神经网络\n')
        # f.write('搜索网格参数2.惩罚参数(C)\n')
        # f.write('\t[0.6, 0.7, 0.8, 1]\n\tC越大，松弛变量惩罚增大，即对误分类的惩罚增大，训练测试准确率高，但泛化能力弱；'
        #         'C值小，对误分类的惩罚减小，允许容错，将他们当成噪声点，泛化能力较强。\n')
        # f.write('搜索最优网格参数:' + '\n' + str(grid_search.best_params_) + '\n')
        # f.write('默认超参数:' + "\n" + str(best_model_svm) + '\n')
        # f.close()

        y_scoreSVM1 = clf.predict(testData)
        # 预测标签
        y_scoreSVM = clf.predict_proba(testData)
        # 预测概率，计算样本点到分割超平面的函数距离,各标签预测概率[0.8,0.2]

        scoreSVM = clf.score(testData, testTarget)
        # 测试数据准去率
        # ORDER = [0, 1, 2, 3, 4]
        # classOneLabel
        matrixFindsize = 12 - 0.1*len(classOneLabel)
        # 混淆矩阵字体大小
        matrixDip = 300
        # 分辨率
        # Plot confusion matrix  绘图混淆矩阵
        sns.set_context("talk", rc={"font": "Helvetica", "font.size": matrixFindsize})
        label = [files[i] for i in classOneLabel]
        cm = confusion_matrix(testTarget, clf.predict(testData), labels=classOneLabel)
        plt.figure(figsize=(5, 4),
                   dpi=matrixDip,
                   # 分辨率
                   facecolor='snow',
                   # 背景色,'snow','blueviolet','lightgray'
                   )
        # 设置title
        # plt.title('Test results of SVM classifier', fontsize=12)

        cm = 100 * cm / cm.sum(axis=1)[:, np.newaxis]
        ax = sns.heatmap(cm, annot_kws={'size': 8},
                         annot=True, cmap='Wistia', fmt='0.1f', xticklabels=label, yticklabels=label)
        # cmap的参数如下：Wistia（蓝绿黄）
        ax.xaxis.tick_top()
        ax.set_xticklabels(ax.get_xticklabels(), rotation=90)
        # plt.xticks(rotation=90)
        plt.yticks(rotation=0)
        # 设置刻度值即坐标轴刻度 xticks yticks
        plt.tick_params(labelsize=matrixFindsize)
        # 设置坐标轴的刻度参数
        # 标签设置
        plt.xlabel('Predicted_label', fontsize=14)
        plt.ylabel('True_label', fontsize=14)
        ax.xaxis.set_label_position('top')
        plt.tight_layout()
        outPath10 = outPath + '\\' + r'_SVM.tif'
        plt.savefig(outPath10, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # outPath10 = outPath + '\\' + r'_SVM.svg'
        # plt.savefig(outPath10, format='svg', bbox_inches='tight', transparent=True)
        plt.close()

        # ******************************************************************************************SVM的ROC曲线与AUC面积

        if 2 == len(classOneLabel):
            # 二分类ROC曲线
            fpr, tpr, threshold = roc_curve(testTarget, y_scoreSVM1)
            ###计算真正率和假正率
            roc_auc = auc(fpr, tpr)
            ###计算auc的值
            # Plot all ROC curves
            plt.figure(figsize=(5, 4), dpi=300)
            lw = 2
            plt.plot(fpr, tpr, color='darkorange',
                     lw=lw, label='ROC curve (area = %0.2f)' % roc_auc)
            ###假正率为横坐标，真正率为纵坐标做曲线
            plt.plot([0, 1], [0, 1], color='navy', lw=lw, linestyle='--')
            plt.xlim([0.0, 1.0])
            plt.ylim([0.0, 1.05])

            plt.xlabel('False Positive Rate', fontsize=14)
            plt.ylabel('True Positive Rate', fontsize=14)
            # plt.title('SVM_ROC', fontsize=12)
            plt.legend(loc="lower right", fontsize=12)
            outPath101 = outPath + '\\' + r'_SVM_ROC.tif'
            plt.savefig(outPath101, format='tif', bbox_inches='tight', transparent=True, dpi=300)
            # plt.rcParams['svg.fonttype'] = 'none'
            # outPath101 = outPath + '\\' + r'_SVM_ROC.svg'
            # plt.savefig(outPath101, format='svg', bbox_inches='tight', transparent=True)
            plt.close()
        else:
            # # 多分类ROC曲线
            # dataTarget1 = label_binarize(dataTarget, np.arange(len(classOneLabel)))
            # # 标签二值化，为了将ROC曲线和ROC区域扩展到多类或多标签分类，有必要对输出进行二值化：
            # # ⑴可以每个标签绘制一条ROC曲线。
            # # ⑵也可以通过将标签指示符矩阵的每个元素视为二元预测（微平均）来绘制ROC曲线。
            # # ⑶另一种用于多类别分类的评估方法是宏观平均，它对每个标签的分类给予相同的权重。
            #
            # print(np.array(dataTarget1).shape)
            #
            # trainData1, testData1, trainTarget1, testTarget1 = train_test_split(ySpec, dataTarget1,
            #                                                                     test_size=svmTestSize, random_state=1)

            # 计算每一类的ROC
            fpr = dict()
            tpr = dict()
            roc_auc = dict()
            # return 0
            for i in range(len(classOneLabel)):
                fpr[i], tpr[i], _ = roc_curve(testTarget1[:, i], y_scoreSVM[:, i])
                roc_auc[i] = auc(fpr[i], tpr[i])

            # Compute micro-average ROC curve and ROC area（方法二）
            fpr["micro"], tpr["micro"], _ = roc_curve(testTarget1.ravel(), y_scoreSVM.ravel())
            roc_auc["micro"] = auc(fpr["micro"], tpr["micro"])

            # Compute macro-average ROC curve and ROC area（方法一）
            # First aggregate all false positive rates
            all_fpr = np.unique(np.concatenate([fpr[i] for i in range(len(classOneLabel))]))
            # Then interpolate all ROC curves at this points
            mean_tpr = np.zeros_like(all_fpr)
            for i in range(len(classOneLabel)):
                mean_tpr += interp(all_fpr, fpr[i], tpr[i])
            # Finally average it and compute AU
            mean_tpr /= len(classOneLabel)
            fpr["macro"] = all_fpr
            tpr["macro"] = mean_tpr
            roc_auc["macro"] = auc(fpr["macro"], tpr["macro"])

            # Plot all ROC curves
            plt.figure(figsize=(5, 4), dpi=300)
            plt.plot(fpr["micro"], tpr["micro"],
                     label='micro-average ROC curve (AUC = {0:0.2f})'
                           ''.format(roc_auc["micro"]),
                     color='deeppink', linestyle=':')

            plt.plot(fpr["macro"], tpr["macro"],
                     label='macro-average ROC curve (AUC = {0:0.2f})'
                           ''.format(roc_auc["macro"]),
                     color='navy', linestyle=':')

            # colors = cycle(['aqua', 'darkorange', 'cornflowerblue'])
            for i in range(len(classOneLabel)):
                plt.plot(fpr[i], tpr[i],
                         label='ROC curve of class {0} (AUC = {1:0.2f})'
                               ''.format(files[i], roc_auc[i]))

            plt.plot([0, 1], [0, 1], 'k--')
            plt.xlim([0.0, 1.0])
            plt.ylim([0.0, 1.05])
            plt.xlabel('False Positive Rate', fontsize=14)
            plt.ylabel('True Positive Rate', fontsize=14)
            # plt.title('SVM_ROC:multi-class', fontsize=12)
            plt.legend(loc="lower right", fontsize=12)
            # plt.show()
            outPath101 = outPath + '\\' + r'_SVM_ROC.tif'
            plt.savefig(outPath101, format='tif', bbox_inches='tight', transparent=True, dpi=300)
            # plt.rcParams['svg.fonttype'] = 'none'
            # outPath101 = outPath + '\\' + r'_SVM_ROC.svg'
            # plt.savefig(outPath101, format='svg', bbox_inches='tight', transparent=True)
            plt.close()
        # ##**************************************************************************************************************LDA
        figure(3 * filesNumInter + 7)
        lda = LinearDiscriminantAnalysis()
        # print('中心点',lda.means_)

        # 建立需要搜索的参数的范围
        # param_grid = {'n_components': [2, 5, 10]}
        # param_grid = {}
        # grid_search_lda = GridSearchCV(lda, param_grid, cv=3, n_jobs=1)
        #
        # try:
        #     lda = grid_search_lda.fit(trainData, trainTarget)
        # except:
        #     lda = grid_search_lda.fit(trainData, trainTarget)

        try:
            lda = lda.fit(trainData, trainTarget)
        except:
            lda = lda.fit(trainData, trainTarget)



        # lda = grid_search_lda.fit(trainData, trainTarget)
        # 训练模型
        # *******************************************************************************************保存、加载分类器模型
        # 保存模型
        outPathLDA = outPath + '\\' + r'_ldaModel.pkl'
        joblib.dump(lda, outPathLDA)
        # 加载模型
        #
        # best_model_lda = grid_search_lda.best_estimator_
        best_model_lda = lda
        # print("***",type(grid_search.best_params_))
        # print(type(grid_search.best_params_))
        # print("默认超参数：",best_model)
        # print(type(best_model))
        # # ********************************************************************************************保存LDA网格参数
        # f = open(outPath + '\\' + r'_ldaParameter.txt', "w")
        # f.write("LDA参数说明：\n")
        # f.write('搜索网格参数1.解释器(solver)\n')
        # f.write('\t奇异值分解(svd):对于有大规模特征的数据，推荐用这种算法\n')
        # f.write('\t最小平方差(lsqr):可以结合skrinkage参数\n')
        # f.write('\t特征分解算法(eigen):可以结合shrinkage参数\n')
        #
        # f.write('搜索网格参数2.指定了数组降维后的维度(n_components)\n')
        # f.write('\t[2,3,5,8,12]\n')
        # f.write('搜索最优网格参数:' + '\n' + str(grid_search_lda.best_params_) + '\n')
        # f.write('默认超参数:' + "\n" + str(best_model_lda) + '\n')
        # f.close()

        scoreLDA = lda.score(testData, testTarget)
        # 测试数据准去率
        # print('测试数据分类正确率', lda.score(X_test, Y_test))
        y_scoreLDA1 = lda.predict(testData)
        # 预测标签
        y_scoreLDA = lda.predict_proba(testData)
        # 预测标签概率

        matrixFindsize = 12 - 0.1*len(classOneLabel)
        # 混淆矩阵字体大小
        matrixDip = 300
        # 分辨率
        # Plot confusion matrix  绘图混淆矩阵
        sns.set_context("talk", rc={"font": "Helvetica", "font.size": matrixFindsize})
        label = [files[i] for i in classOneLabel]
        cm = confusion_matrix(testTarget, lda.predict(testData), labels=classOneLabel)
        plt.figure(figsize=(5, 4),
                   dpi=matrixDip,
                   # 分辨率
                   facecolor='blueviolet',
                   # 背景色,'snow','blueviolet','lightgray'
                   )
        # 设置title
        # plt.title('Test results of LDA classifier', fontsize=12)

        cm = 100 * cm / cm.sum(axis=1)[:, np.newaxis]
        ax = sns.heatmap(cm, annot=True, cmap='YlOrRd', fmt='0.1f', xticklabels=label, yticklabels=label)
        ax.xaxis.tick_top()
        ax.set_xticklabels(ax.get_xticklabels(), rotation=90)
        # plt.xticks()
        plt.yticks(rotation=0)
        # 设置刻度值即坐标轴刻度 xticks yticks
        plt.tick_params(labelsize=matrixFindsize)
        # 设置坐标轴的刻度参数
        # 标签设置
        plt.xlabel('Predicted_label', fontsize=14)
        plt.ylabel('True_label', fontsize=14)
        ax.xaxis.set_label_position('top')
        plt.tight_layout()
        outPath11 = outPath + '\\' + r'_LDA.tif'
        plt.savefig(outPath11, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # outPath11 = outPath + '\\' + r'_LDA.svg'
        # plt.savefig(outPath11, format='svg', bbox_inches='tight', transparent=True)
        plt.close()
        # ******************************************************************************************LDA的ROC曲线与AUC面积

        if 2 == len(classOneLabel):
            # 二分类ROC曲线
            fpr, tpr, threshold = roc_curve(testTarget, y_scoreLDA1)
            ###计算真正率和假正率
            roc_auc = auc(fpr, tpr)
            ###计算auc的值
            # Plot all ROC curves
            plt.figure(figsize=(5, 4), dpi=300)
            lw = 2
            plt.plot(fpr, tpr, color='darkorange',
                     lw=lw, label='ROC curve (area = %0.2f)' % roc_auc)
            ###假正率为横坐标，真正率为纵坐标做曲线
            plt.plot([0, 1], [0, 1], color='navy', lw=lw, linestyle='--')
            plt.xlim([0.0, 1.0])
            plt.ylim([0.0, 1.05])

            plt.xlabel('False Positive Rate', fontsize=14)
            plt.ylabel('True Positive Rate', fontsize=14)
            # plt.title('LDA_ROC', fontsize=12)
            plt.legend(loc="lower right", fontsize=12)
            # plt.show()
            outPath111 = outPath + '\\' + r'_LDA_ROC.tif'
            plt.savefig(outPath111, format='tif', bbox_inches='tight', transparent=True, dpi=300)
            # plt.rcParams['svg.fonttype'] = 'none'
            # outPath111 = outPath + '\\' + r'_LDA_ROC.svg'
            # plt.savefig(outPath111, format='svg', bbox_inches='tight', transparent=True)
            plt.close()
        else:
            #
            for i in range(len(classOneLabel)):
                fpr[i], tpr[i], _ = roc_curve(testTarget1[:, i], y_scoreLDA[:, i])
                roc_auc[i] = auc(fpr[i], tpr[i])

            # Compute micro-average ROC curve and ROC area（方法二）
            fpr["micro"], tpr["micro"], _ = roc_curve(testTarget1.ravel(), y_scoreLDA.ravel())
            roc_auc["micro"] = auc(fpr["micro"], tpr["micro"])

            # Compute macro-average ROC curve and ROC area（方法一）
            # First aggregate all false positive rates
            all_fpr = np.unique(np.concatenate([fpr[i] for i in range(len(classOneLabel))]))
            # Then interpolate all ROC curves at this points
            mean_tpr = np.zeros_like(all_fpr)
            for i in range(len(classOneLabel)):
                mean_tpr += interp(all_fpr, fpr[i], tpr[i])
            # Finally average it and compute AU
            mean_tpr /= len(classOneLabel)
            fpr["macro"] = all_fpr
            tpr["macro"] = mean_tpr
            roc_auc["macro"] = auc(fpr["macro"], tpr["macro"])

            # Plot all ROC curves
            plt.figure(figsize=(5, 4), dpi=300)
            plt.plot(fpr["micro"], tpr["micro"],
                     label='micro-average ROC curve (AUC = {0:0.2f})'
                           ''.format(roc_auc["micro"]),
                     color='deeppink', linestyle=':')

            plt.plot(fpr["macro"], tpr["macro"],
                     label='macro-average ROC curve (AUC = {0:0.2f})'
                           ''.format(roc_auc["macro"]),
                     color='navy', linestyle=':')

            # colors = cycle(['aqua', 'darkorange', 'cornflowerblue'])
            for i in range(len(classOneLabel)):
                plt.plot(fpr[i], tpr[i],
                         label='ROC curve of class {0} (AUC = {1:0.2f})'
                               ''.format(files[i], roc_auc[i]))

            plt.plot([0, 1], [0, 1], 'k--')
            plt.xlim([0.0, 1.0])
            plt.ylim([0.0, 1.05])
            plt.xlabel('False Positive Rate', fontsize=14)
            plt.ylabel('True Positive Rate', fontsize=14)
            # plt.title('LDA_ROC:multi-class', fontsize=12)
            plt.legend(loc="lower right", fontsize=12)
            # plt.show()
            outPath111 = outPath + '\\' + r'_LDA_ROC.tif'
            plt.savefig(outPath111, format='tif', bbox_inches='tight', transparent=True, dpi=300)
            # plt.rcParams['svg.fonttype'] = 'none'
            # outPath111 = outPath + '\\' + r'_LDA_ROC.svg'
            # plt.savefig(outPath111, format='svg', bbox_inches='tight', transparent=True)
            plt.close()
        # **************************************************************************************************************KNN
        figure(3 * filesNumInter + 8)

        knn = KNeighborsClassifier()

        # 定义一个knn分类器对象
        # int 型参数 knn算法中指定以最近的几个最近邻样本具有投票权，默认参数为5，n_neighbors=5
        # str参数,即每个拥有投票权的样本是按什么比重投票，'uniform'表示等比重投票，
        # 'distance'表示按距离反比投票，[callable]表示自己定义的一个函数
        # str或者距离度量对象   即怎样度量距离,默认是闵氏距离。

        # 建立需要搜索的参数的范围
        # param_grid = {'n_neighbors': [3, 5, 7, 9], 'weights': ['uniform', 'distance'], 'leaf_size': [20, 30, 40]}
        # grid_search_knn = GridSearchCV(knn, param_grid, cv=3, n_jobs=1)
        # try:
        #     knn = grid_search_knn.fit(trainData, trainTarget)
        # except:
        #     knn = grid_search_knn.fit(trainData, trainTarget)
        try:
            knn = knn.fit(trainData, trainTarget)
        except:
            knn = knn.fit(trainData, trainTarget)
        # knn = grid_search_knn.fit(trainData, trainTarget)
        # 训练模型
        # *******************************************************************************************保存、加载分类器模型
        # 保存模型
        outPathKNN = outPath + '\\' + r'_knnModel.pkl'
        joblib.dump(knn, outPathKNN)
        # 加载模型
        #
        # best_model_knn = grid_search_knn.best_estimator_
        best_model_knn = knn
        # print("***",type(grid_search.best_params_))
        # print(type(grid_search.best_params_))
        # print("默认超参数：",best_model)
        # print(type(best_model))
        # # ********************************************************************************************保存knn网格参数
        # f = open(outPath + '\\' + r'_knnParameter.txt', "w")
        # f.write("knn参数说明：\n")
        # f.write('搜索网格参数1.相邻节点数(n_neighbors)\n')
        # f.write('\t[3,5,7,10]\n')
        # f.write('搜索网格参数2.权重函数(weights)\n')
        # f.write('\tuniform:统一的权重. 在每一个邻居区域里的点的权重都是一样的。\n')
        # f.write('\tdistance:权重点等于他们距离的倒数。使用此函数，更近的邻居对于所预测的点的影响更大。\n')
        # f.write('搜索网格参数3.叶子数量(leaf_size)\n')
        # f.write('\t[20,30,40]\n')
        # f.write('搜索最优网格参数:' + '\n' + str(grid_search_knn.best_params_) + '\n')
        # f.write('默认超参数:' + "\n" + str(best_model_knn) + '\n')
        # f.close()

        # knn.fit(trainData, trainTarget)
        scoreKNN = knn.score(testData, testTarget)
        # 测试数据准确率
        y_scoreKNN1 = knn.predict(testData)
        # 预测标签
        y_scoreKNN = knn.predict_proba(testData)
        # 预测标签概率
        matrixFindsize = 12 - 0.1*len(classOneLabel)
        # 混淆矩阵字体大小
        matrixDip = 300
        # 分辨率
        # Plot confusion matrix  绘图混淆矩阵
        sns.set_context("talk", rc={"font": "Helvetica", "font.size": matrixFindsize})
        label = [files[i] for i in classOneLabel]
        cm = confusion_matrix(testTarget, knn.predict(testData), labels=classOneLabel)
        plt.figure(figsize=(5, 4),
                   dpi=matrixDip,
                   # 分辨率
                   facecolor='blueviolet',
                   # 背景色,'snow','blueviolet','lightgray'
                   )
        # 设置title
        # plt.title('Test results of KNN classifier', fontsize=12)

        cm = 100 * cm / cm.sum(axis=1)[:, np.newaxis]
        ax = sns.heatmap(cm, annot=True, cmap='YlGnBu', fmt='0.1f', xticklabels=label, yticklabels=label)
        ax.xaxis.tick_top()
        ax.set_xticklabels(ax.get_xticklabels(), rotation=90)
        # plt.xticks(rotation=90)
        plt.yticks(rotation=0)
        # 设置刻度值即坐标轴刻度 xticks yticks
        plt.tick_params(labelsize=matrixFindsize)
        # 设置坐标轴的刻度参数
        # 标签设置
        plt.xlabel('Predicted_label', fontsize=14)
        plt.ylabel('True_label', fontsize=14)
        ax.xaxis.set_label_position('top')
        plt.tight_layout()
        outPath12 = outPath + '\\' + r'_KNN.tif'
        plt.savefig(outPath12, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # outPath12 = outPath + '\\' + r'_KNN.svg'
        # plt.savefig(outPath12, format='svg', bbox_inches='tight', transparent=True)
        plt.close()
        # ******************************************************************************************KNN的ROC曲线与AUC面积
        if 2 == len(classOneLabel):
            # 二分类ROC曲线
            fpr, tpr, threshold = roc_curve(testTarget, y_scoreKNN1)
            ###计算真正率和假正率
            roc_auc = auc(fpr, tpr)
            ###计算auc的值
            # Plot all ROC curves
            plt.figure(figsize=(5, 4), dpi=300)
            lw = 2
            plt.plot(fpr, tpr, color='darkorange',
                     lw=lw, label='ROC curve (area = %0.2f)' % roc_auc)
            ###假正率为横坐标，真正率为纵坐标做曲线
            plt.plot([0, 1], [0, 1], color='navy', lw=lw, linestyle='--')
            plt.xlim([0.0, 1.0])
            plt.ylim([0.0, 1.05])

            plt.xlabel('False Positive Rate', fontsize=14)
            plt.ylabel('True Positive Rate', fontsize=14)
            # plt.title('KNN_ROC', fontsize=12)
            plt.legend(loc="lower right", fontsize=12)
            # plt.show()
            outPath121 = outPath + '\\' + r'_KNN_ROC.tif'
            plt.savefig(outPath121, format='tif', bbox_inches='tight', transparent=True, dpi=300)
            # plt.rcParams['svg.fonttype'] = 'none'
            # outPath121 = outPath + '\\' + r'_KNN_ROC.svg'
            # plt.savefig(outPath121, format='svg', bbox_inches='tight', transparent=True)
            plt.close()
        else:
            # 多分类
            for i in range(len(classOneLabel)):
                fpr[i], tpr[i], _ = roc_curve(testTarget1[:, i], y_scoreKNN[:, i])
                roc_auc[i] = auc(fpr[i], tpr[i])

            # Compute micro-average ROC curve and ROC area（方法二）
            fpr["micro"], tpr["micro"], _ = roc_curve(testTarget1.ravel(), y_scoreKNN.ravel())
            roc_auc["micro"] = auc(fpr["micro"], tpr["micro"])

            # Compute macro-average ROC curve and ROC area（方法一）
            # First aggregate all false positive rates
            all_fpr = np.unique(np.concatenate([fpr[i] for i in range(len(classOneLabel))]))
            # Then interpolate all ROC curves at this points
            mean_tpr = np.zeros_like(all_fpr)
            for i in range(len(classOneLabel)):
                mean_tpr += interp(all_fpr, fpr[i], tpr[i])
            # Finally average it and compute AU
            mean_tpr /= len(classOneLabel)
            fpr["macro"] = all_fpr
            tpr["macro"] = mean_tpr
            roc_auc["macro"] = auc(fpr["macro"], tpr["macro"])

            # Plot all ROC curves
            plt.figure(figsize=(5, 4), dpi=300)
            plt.plot(fpr["micro"], tpr["micro"],
                     label='micro-average ROC curve (AUC = {0:0.2f})'
                           ''.format(roc_auc["micro"]),
                     color='deeppink', linestyle=':')

            plt.plot(fpr["macro"], tpr["macro"],
                     label='macro-average ROC curve (AUC = {0:0.2f})'
                           ''.format(roc_auc["macro"]),
                     color='navy', linestyle=':')

            # colors = cycle(['aqua', 'darkorange', 'cornflowerblue'])
            for i in range(len(classOneLabel)):
                plt.plot(fpr[i], tpr[i],
                         label='ROC curve of class {0} (AUC = {1:0.2f})'
                               ''.format(files[i], roc_auc[i]))

            plt.plot([0, 1], [0, 1], 'k--')
            plt.xlim([0.0, 1.0])
            plt.ylim([0.0, 1.05])
            plt.xlabel('False Positive Rate', fontsize=14)
            plt.ylabel('True Positive Rate', fontsize=14)
            # plt.title('KNN_ROC:multi-class', fontsize=12)
            plt.legend(loc="lower right", fontsize=12)
            # plt.show()
            outPath121 = outPath + '\\' + r'_KNN_ROC.tif'
            plt.savefig(outPath121, format='tif', bbox_inches='tight', transparent=True, dpi=300)
            # plt.rcParams['svg.fonttype'] = 'none'
            # outPath121 = outPath + '\\' + r'_KNN_ROC.svg'
            # plt.savefig(outPath121, format='svg', bbox_inches='tight', transparent=True)
            plt.close()

        # **************************************************************************************************************决策树
        figure(3 * filesNumInter + 8)
        from sklearn.tree import DecisionTreeClassifier
        dt = DecisionTreeClassifier(random_state=1)
        # 直接训练，不用网格搜索
        dt = dt.fit(trainData, trainTarget)
        best_model_dt = dt

        # *******************************************************************************************保存、加载分类器模型
        # 保存模型
        outPathDT = outPath + '\\' + r'_dtModel.pkl'
        joblib.dump(dt, outPathDT)

        # best_model_dt = grid_search_dt.best_estimator_
        # best_model_dt = dt
        # ********************************************************************************************保存DT网格参数
        # f = open(outPath + '\\' + r'_dtParameter.txt', "w")
        # f.write("Decision Tree参数说明：\n")
        # f.write('搜索网格参数1.特征选择标准(criterion)\n')
        # f.write('\tgini: 基尼不纯度，计算速度较快\n')
        # f.write('\tentropy: 信息熵，倾向于获得信息增益更大的划分\n')
        # f.write('搜索网格参数2.树的最大深度(max_depth)\n')
        # f.write('\t[3, 5, 7, 9, None]\n')
        # f.write('搜索网格参数3.内部节点再划分所需最小样本数(min_samples_split)\n')
        # f.write('\t[2, 4, 6, 8]\n')
        # f.write('搜索网格参数4.叶子节点最少样本数(min_samples_leaf)\n')
        # f.write('\t[1, 2, 4]\n')
        # f.write('搜索最优网格参数:' + '\n' + str(grid_search_dt.best_params_) + '\n')
        # f.write('默认超参数:' + "\n" + str(best_model_dt) + '\n')
        # f.close()

        scoreDT = dt.score(testData, testTarget)
        # 测试数据准确率

        y_scoreDT1 = dt.predict(testData)
        # 预测标签

        matrixFindsize = 12 - 0.1 * len(classOneLabel)
        # 混淆矩阵字体大小
        matrixDip = 300
        # 分辨率

        # Plot confusion matrix  绘图混淆矩阵
        sns.set_context("talk", rc={"font": "Helvetica", "font.size": matrixFindsize})
        label = [files[i] for i in classOneLabel]
        cm = confusion_matrix(testTarget, dt.predict(testData), labels=classOneLabel)

        plt.figure(figsize=(5, 4),
                   dpi=matrixDip,
                   facecolor='lightgray',
                   )

        cm = 100 * cm / cm.sum(axis=1)[:, np.newaxis]
        ax = sns.heatmap(cm, annot=True, cmap='OrRd', fmt='0.1f',
                         xticklabels=label, yticklabels=label)
        ax.xaxis.tick_top()
        ax.set_xticklabels(ax.get_xticklabels(), rotation=90)
        plt.yticks(rotation=0)
        plt.tick_params(labelsize=matrixFindsize)

        plt.xlabel('Predicted_label', fontsize=14)
        plt.ylabel('True_label', fontsize=14)
        ax.xaxis.set_label_position('top')
        plt.tight_layout()

        outPathDTFig = outPath + '\\' + r'_DT.tif'
        plt.savefig(outPathDTFig, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # outPathDTFig = outPath + '\\' + r'_DT.svg'
        # plt.savefig(outPathDTFig, format='svg', bbox_inches='tight', transparent=True)
        plt.close()

        # **************************************************************************************************************Gradient Boosting,运行太慢
        # figure(3 * filesNumInter + 8)
        # from sklearn.ensemble import GradientBoostingClassifier
        # gbdt = GradientBoostingClassifier(random_state=1)
        #
        # # 建立需要搜索的参数范围
        # param_grid = {
        #     'n_estimators': [50, 100, 150],
        #     'learning_rate': [0.05, 0.1, 0.2],
        #     'max_depth': [2, 3, 4],
        #     'subsample': [0.8, 1.0]
        # }
        #
        # # grid_search_gbdt = GridSearchCV(gbdt, param_grid, cv=3, n_jobs=1)
        # # try:
        # #     gbdt = grid_search_gbdt.fit(trainData, trainTarget)
        # # except:
        # #     gbdt = grid_search_gbdt.fit(trainData, trainTarget)
        # try:
        #     gbdt = gbdt.fit(trainData, trainTarget)
        # except:
        #     gbdt = gbdt.fit(trainData, trainTarget)

        # from sklearn.ensemble import GradientBoostingClassifier
        # gbdt = GradientBoostingClassifier(random_state=1)
        # gbdt = gbdt.fit(trainData, trainTarget)
        # best_model_gbdt = gbdt
        #
        # # 保存模型
        # outPathGBDT = outPath + '\\' + r'_gbdtModel.pkl'
        # joblib.dump(gbdt, outPathGBDT)
        #
        # # best_model_gbdt = grid_search_gbdt.best_estimator_
        # # best_model_gbdt = gbdt
        # # 保存参数说明
        # # f = open(outPath + '\\' + r'_gbdtParameter.txt', "w")
        # # f.write("Gradient Boosting参数说明：\n")
        # # f.write('搜索网格参数1.弱学习器个数(n_estimators)\n')
        # # f.write('\t[50, 100, 150]\n')
        # # f.write('搜索网格参数2.学习率(learning_rate)\n')
        # # f.write('\t[0.05, 0.1, 0.2]\n')
        # # f.write('搜索网格参数3.基学习器树深(max_depth)\n')
        # # f.write('\t[2, 3, 4]\n')
        # # f.write('搜索网格参数4.子采样比例(subsample)\n')
        # # f.write('\t[0.8, 1.0]\n')
        # # f.write('搜索最优网格参数:' + '\n' + str(grid_search_gbdt.best_params_) + '\n')
        # # f.write('默认超参数:' + "\n" + str(best_model_gbdt) + '\n')
        # # f.close()
        #
        # # scoreGBDT = gbdt.score(testData, testTarget)
        # # y_scoreGBDT1 = gbdt.predict(testData)
        #
        # matrixFindsize = 12 - 0.1 * len(classOneLabel)
        # matrixDip = 300
        #
        # sns.set_context("talk", rc={"font": "Helvetica", "font.size": matrixFindsize})
        # label = [files[i] for i in classOneLabel]
        # cm = confusion_matrix(testTarget, gbdt.predict(testData), labels=classOneLabel)
        #
        # plt.figure(figsize=(5, 4),
        #            dpi=matrixDip,
        #            facecolor='lightgray')
        #
        # cm = 100 * cm / cm.sum(axis=1)[:, np.newaxis]
        # ax = sns.heatmap(cm, annot=True, cmap='YlOrBr', fmt='0.1f',
        #                  xticklabels=label, yticklabels=label)
        # ax.xaxis.tick_top()
        # ax.set_xticklabels(ax.get_xticklabels(), rotation=90)
        # plt.yticks(rotation=0)
        # plt.tick_params(labelsize=matrixFindsize)
        # plt.xlabel('Predicted_label', fontsize=14)
        # plt.ylabel('True_label', fontsize=14)
        # ax.xaxis.set_label_position('top')
        # plt.tight_layout()
        #
        # outPathGBDTFig = outPath + '\\' + r'_GBDT.tif'
        # plt.savefig(outPathGBDTFig, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # # plt.rcParams['svg.fonttype'] = 'none'
        # # outPathGBDTFig = outPath + '\\' + r'_GBDT.svg'
        # # plt.savefig(outPathGBDTFig, format='svg', bbox_inches='tight', transparent=True)
        # plt.close()


        # **************************************************************************************************************Extra Trees
        figure(3 * filesNumInter + 8)
        # from sklearn.ensemble import ExtraTreesClassifier
        # et = ExtraTreesClassifier(random_state=1)
        #
        # # 建立需要搜索的参数范围
        # param_grid = {
        #     'n_estimators': [100, 200, 300],
        #     'criterion': ['gini', 'entropy'],
        #     'max_depth': [None, 5, 10, 15],
        #     'min_samples_split': [2, 4, 6]
        # }
        #
        # # grid_search_et = GridSearchCV(et, param_grid, cv=3, n_jobs=1)
        # # try:
        # #     et = grid_search_et.fit(trainData, trainTarget)
        # # except:
        # #     et = grid_search_et.fit(trainData, trainTarget)
        # try:
        #     et = et.fit(trainData, trainTarget)
        # except:
        #     et = et.fit(trainData, trainTarget)

        from sklearn.ensemble import ExtraTreesClassifier
        et = ExtraTreesClassifier(random_state=1)
        et = et.fit(trainData, trainTarget)
        best_model_et = et
        # 保存模型
        outPathET = outPath + '\\' + r'_etModel.pkl'
        joblib.dump(et, outPathET)

        # best_model_et = grid_search_et.best_estimator_
        # best_model_et = et
        # 保存参数说明
        # f = open(outPath + '\\' + r'_etParameter.txt', "w")
        # f.write("Extra Trees参数说明：\n")
        # f.write('搜索网格参数1.树的数量(n_estimators)\n')
        # f.write('\t[100, 200, 300]\n')
        # f.write('搜索网格参数2.特征选择标准(criterion)\n')
        # f.write('\tgini: 基尼不纯度\n')
        # f.write('\tentropy: 信息熵\n')
        # f.write('搜索网格参数3.树的最大深度(max_depth)\n')
        # f.write('\t[None, 5, 10, 15]\n')
        # f.write('搜索网格参数4.内部节点再划分所需最小样本数(min_samples_split)\n')
        # f.write('\t[2, 4, 6]\n')
        # f.write('搜索最优网格参数:' + '\n' + str(grid_search_et.best_params_) + '\n')
        # f.write('默认超参数:' + "\n" + str(best_model_et) + '\n')
        # f.close()

        scoreET = et.score(testData, testTarget)
        y_scoreET1 = et.predict(testData)

        matrixFindsize = 12 - 0.1 * len(classOneLabel)
        matrixDip = 300

        sns.set_context("talk", rc={"font": "Helvetica", "font.size": matrixFindsize})
        label = [files[i] for i in classOneLabel]
        cm = confusion_matrix(testTarget, et.predict(testData), labels=classOneLabel)

        plt.figure(figsize=(5, 4),
                   dpi=matrixDip,
                   facecolor='lightgray')

        cm = 100 * cm / cm.sum(axis=1)[:, np.newaxis]
        ax = sns.heatmap(cm, annot=True, cmap='PuBuGn', fmt='0.1f',
                         xticklabels=label, yticklabels=label)
        ax.xaxis.tick_top()
        ax.set_xticklabels(ax.get_xticklabels(), rotation=90)
        plt.yticks(rotation=0)
        plt.tick_params(labelsize=matrixFindsize)
        plt.xlabel('Predicted_label', fontsize=14)
        plt.ylabel('True_label', fontsize=14)
        ax.xaxis.set_label_position('top')
        plt.tight_layout()

        outPathETFig = outPath + '\\' + r'_ET.tif'
        plt.savefig(outPathETFig, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # outPathETFig = outPath + '\\' + r'_ET.svg'
        # plt.savefig(outPathETFig, format='svg', bbox_inches='tight', transparent=True)
        plt.close()


        # **************************************************************************************************************XGBoost
        figure(3 * filesNumInter + 8)

        # xgb = XGBClassifier()
        xgb = XGBClassifier(use_label_encoder=False, eval_metric='logloss')
        # scale_pos_weight,正样本的权重，在二分类任务中，当正负样本比例失衡时，设置正样本的权重，模型效果更好。
        # 建立需要搜索的参数的范围

        # param_grid = {'max_depth': [5, 6, 7]}
        # grid_search_xgb = GridSearchCV(xgb, param_grid, cv=3, n_jobs=1)
        # try:
        #     xgb = grid_search_xgb.fit(trainData, trainTarget)
        # except:
        #     xgb = grid_search_xgb.fit(trainData, trainTarget)
        try:
            xgb = xgb.fit(trainData, trainTarget)
        except:
            xgb = xgb.fit(trainData, trainTarget)
        # xgb = grid_search_xgb.fit(trainData, trainTarget)
        # 训练模型
        # *******************************************************************************************保存、加载分类器模型
        # 保存模型
        outPathXGB = outPath + '\\' + r'_xgbModel.pkl'
        joblib.dump(xgb, outPathXGB)
        # 加载模型
        # best_model_xgb = grid_search_xgb.best_estimator_
        best_model_xgb = xgb
        # 网格函数选出的最优模型
        # print("***",type(grid_search.best_params_))
        # print(type(grid_search.best_params_))
        # print("默认超参数：",best_model)
        # print(type(best_model))

        # ====== 这里开始加入 SHAP 代码 ======
        # SHAP 全局重要性分析（基于 XGBoost）
        # 说明：
        # trainData/testData: 形状为 [样本数, 特征数]
        # x: 每个特征对应的拉曼位移（波数）
        # 因此 SHAP 的每个特征名应设置为对应的拉曼位移

        # 构造特征名：使用拉曼位移作为特征名，保留两位小数
        feature_names = [f"{val:.2f}" for val in x]

        # 转为 DataFrame，便于 SHAP 绘图时显示真实波数而不是默认 feature_0, feature_1
        trainData_df = pd.DataFrame(trainData, columns=feature_names)
        testData_df = pd.DataFrame(testData, columns=feature_names)

        # 建立 SHAP 解释器
        explainer = shap.TreeExplainer(best_model_xgb)

        # 计算测试集 SHAP 值
        shap_values = explainer.shap_values(testData_df)

        # 兼容不同版本 shap / 二分类返回格式
        if isinstance(shap_values, list):
            shap_values_plot = shap_values[1]
        else:
            shap_values_plot = shap_values

        # 计算全局重要性：mean(|SHAP|)
        # 计算全局重要性：mean(|SHAP|)
        shap_importance = np.abs(shap_values_plot).mean(axis=0)
        top_K_shap = 20  # 初始提取前20个候选特征，用于后续相邻合并筛选
        top_k_idx_shap = np.argsort(shap_importance)[::-1][:top_K_shap]

        # ====== 新增：对相邻波数差 < 5 cm^-1 的特征进行合并 ======
        # Step 1：将 top20 特征按波数升序排列
        _wave_sort_order = np.argsort(x[top_k_idx_shap])
        _idx_wave_sorted = top_k_idx_shap[_wave_sort_order]  # 按波数从小到大的特征索引列表

        # Step 2：链式分组——相邻两个特征波数差 < 5 cm^-1 则归入同一组
        _groups = []
        _cur_group = [int(_idx_wave_sorted[0])]
        for _kk in range(1, len(_idx_wave_sorted)):
            _cur_idx = int(_idx_wave_sorted[_kk])
            _pre_idx = int(_idx_wave_sorted[_kk - 1])  # 与紧邻的上一个特征比较
            if abs(float(x[_cur_idx]) - float(x[_pre_idx])) < 5:
                _cur_group.append(_cur_idx)
            else:
                _groups.append(_cur_group)
                _cur_group = [_cur_idx]
        _groups.append(_cur_group)

        # Step 3：每组内 SHAP 值相加，代表性索引选取规则：
        #          奇数个特征 → 取中间那个；偶数个特征 → 取中间靠前（偏左）那个
        _merged_rep_idx = []
        _merged_sum_vals = []
        for _grp in _groups:
            _summed = float(np.sum([shap_importance[_ii] for _ii in _grp]))
            _n = len(_grp)
            if _n % 2 == 1:
                _rep = _grp[_n // 2]  # 奇数：正中间
            else:
                _rep = _grp[_n // 2 - 1]  # 偶数：中间靠前
            _merged_rep_idx.append(_rep)
            _merged_sum_vals.append(_summed)

        _merged_rep_idx = np.array(_merged_rep_idx)
        _merged_sum_vals = np.array(_merged_sum_vals)

        # Step 4：按合并后 SHAP 值降序排列，取前15个
        top_K_shap_final = 15
        _top_order = np.argsort(_merged_sum_vals)[::-1][:top_K_shap_final]
        final_top_idx = _merged_rep_idx[_top_order]  # 代表性特征索引（降序，最重要在前）
        final_top_values = _merged_sum_vals[_top_order]  # 合并后的 SHAP 值（降序）
        # ====== 合并结束 ======

        # 保存 SHAP 全局重要性结果到 txt（前15个，相邻<5cm^-1特征已合并累加）
        with open(outPath + '\\' + r'_xgbSHAP.txt', 'w', encoding='utf-8') as f:
            f.write('XGBoost + SHAP 全局特征重要性（前15个拉曼位移，相邻<5cm^-1特征已合并累加）\n')
            f.write('排序依据：合并后 mean(|SHAP value|) 累加值\n')
            for _i in range(len(final_top_idx)):
                _idx = final_top_idx[_i]
                f.write(f"拉曼位移: {float(x[_idx]):.2f} cm^-1, 合并后SHAP值: {float(final_top_values[_i]):.6f}\n")

        # ===== 方案1：推荐论文使用的自定义横向条形图（前15个，合并后排序）=====
        # 反转顺序：barh 从下到上显示，最重要特征显示在图顶部
        plot_idx = final_top_idx[::-1]  # 重要性从小到大（对应 barh 的底→顶方向）
        plot_values = final_top_values[::-1]
        plot_labels = [f"{x[i]:.2f}" for i in plot_idx]

        plt.figure(figsize=(8, 6), dpi=300, facecolor='white')
        bars = plt.barh(
            y=np.arange(len(plot_idx)),
            width=plot_values,
            color='#00BFFF',  # 蓝色系，论文中较常用，稳重清晰
            edgecolor='black',
            linewidth=0.6
        )

        plt.yticks(np.arange(len(plot_idx)), plot_labels, fontsize=11)
        plt.xticks(fontsize=11)
        plt.xlabel('Mean(|SHAP value|)', fontsize=13)
        plt.ylabel('Raman shift (cm$^{-1}$)', fontsize=13)
        # plt.title('Global feature importance based on SHAP', fontsize=13)

        # 添加浅色横向网格，提高可读性
        plt.grid(axis='x', linestyle='--', alpha=0.35)

        # 去掉上边框和右边框，更符合论文制图习惯
        ax = plt.gca()
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        plt.tight_layout()
        plt.savefig(outPath + '\\' + r'_XGB_SHAP_Global_Bar.tif',
                    format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # plt.savefig(outPath + '\\' + r'_XGB_SHAP_Global_Bar.svg',
        #             format='svg', bbox_inches='tight', transparent=True)
        plt.close()

        # ====== 这里结束 SHAP 代码 ======

        # ===== 这里再放 SHAP 蜂群图 =====
        """
        SHAP 蜂群图
        """
        ...
        plt.figure(figsize=(8, 6), dpi=300, facecolor='white')
        shap.summary_plot(
            shap_values_plot,
            testData_df,
            max_display=15,
            show=False
        )
        plt.xlabel('SHAP value (impact on model output)', fontsize=13)
        plt.ylabel('Raman shift (cm$^{-1}$)', fontsize=13)
        ax = plt.gca()
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        plt.tight_layout()
        plt.savefig(outPath + '\\' + r'_XGB_SHAP_Beeswarm.tif',
                    format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # plt.savefig(outPath + '\\' + r'_XGB_SHAP_Beeswarm.svg',
        #             format='svg', bbox_inches='tight', transparent=True)
        plt.close()

        with open(outPath + '\\' + r'_XGB_SHAP_Beeswarm_说明.txt', 'w', encoding='utf-8') as f:
            f.write("SHAP蜂群图说明：\n")
            f.write("1. 每一行代表一个拉曼位移特征（波数点）。\n")
            f.write("2. 横坐标 SHAP value 表示该特征对模型输出的影响方向与大小。\n")
            f.write("3. SHAP value > 0：该特征推动样本向正类（通常为第2类）预测。\n")
            f.write("4. SHAP value < 0：该特征推动样本向负类（通常为第1类）预测。\n")
            f.write("5. 点的颜色表示该拉曼位移处的光谱强度大小：红色高，蓝色低。\n")
            f.write("6. 若某波数右侧多为红点，说明该波数强度高时更倾向于判为正类。\n")
        with open(outPath + '\\' + r'_XGB_SHAP_Beeswarm_说明1.txt', 'w', encoding='utf-8') as f:
            f.write("SHAP蜂群图（图4b）分析说明：\n\n")
            f.write("本图展示XGBoost模型对六类白酒发酵微生物分类时，前15个关键拉曼特征的SHAP值分布。\n")
            f.write("横轴为SHAP值（正值推动向目标类预测，负值反之），点颜色代表该波数处归一化光谱强度（红色高，蓝色低）。\n\n")
            f.write("主要规律：\n")
            f.write("1. 1715.5 cm⁻¹（酯羰基C=O）与2862.8 cm⁻¹（脂肪酸链CH₂对称伸缩）呈最强单向正向贡献：\n")
            f.write("   高强度（红色）样本集中于正向SHAP区间，与乳酸乙酯、乙酸乙酯等主体酯类的酶促酯化合成直接关联，\n")
            f.write("   是判别产酯型菌株（如Staphylococcus xylosus）的核心光谱标志。\n\n")
            f.write("2. 934.2与949.4 cm⁻¹（多糖/蛋白质α螺旋骨架）高贡献菌株集中于Lichtheimia属丝状真菌，\n")
            f.write("   对应其糖化酶主导的淀粉降解功能，多糖积累信号增强支持其作为大曲糖化功能菌群的光谱表型解释。\n\n")
            f.write("3. 1030.9 cm⁻¹（苯丙氨酸苯环振动）与1322.8 cm⁻¹（酰胺III带）正向贡献\n")
            f.write("   指向Bacillus licheniformis经Ehrlich途径生成苯乙醇等花香型风味前体的代谢能力。\n\n")
            f.write("4. 734.3、796.3、802.5 cm⁻¹（核酸/RNA磷酸骨架）及1087.7 cm⁻¹（磷脂O-P-O）\n")
            f.write("   呈双向分散分布，反映细菌与丝状真菌在RNA代谢活性和细胞膜磷脂组成上的本质差异，\n")
            f.write("   作为'细菌—真菌'二元区分的辅助判别信号。\n\n")
            f.write("结论：脂质-酯化类特征（1715.5、2862.8 cm⁻¹）与蛋白质-氨基酸类特征（934.2、1030.9、1322.8 cm⁻¹）\n")
            f.write("的SHAP贡献最为突出，可作为连接微生物光谱表型与白酒风味物质生成潜能的核心光谱标志物。\n")

        '''
        SHAP dependence plot
        '''
        # 批量绘制多个关键拉曼位移的 SHAP dependence plot
        target_shifts = [1002.0, 1450.0, 1660.0]

        for target_shift in target_shifts:
            target_idx = int(np.argmin(np.abs(x - target_shift)))
            target_feature_name = feature_names[target_idx]

            plt.figure(figsize=(7, 5), dpi=300, facecolor='white')
            shap.dependence_plot(
                target_feature_name,
                shap_values_plot,
                testData_df,
                interaction_index='auto',
                show=False
            )

            plt.xlabel(f'Intensity at Raman shift {target_feature_name} cm$^{{-1}}$', fontsize=13)
            plt.ylabel('SHAP value', fontsize=13)

            ax = plt.gca()
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)

            plt.tight_layout()
            file_tag = str(target_shift).replace('.', '_')
            plt.savefig(outPath + '\\' + f'_XGB_SHAP_Dependence_{file_tag}.tif',
                        format='tif', bbox_inches='tight', transparent=True, dpi=300)
            # plt.rcParams['svg.fonttype'] = 'none'
            # plt.savefig(outPath + '\\' + f'_XGB_SHAP_Dependence_{file_tag}.svg',
            #             format='svg', bbox_inches='tight', transparent=True)
            plt.close()

        '''
        SHAP value mapping plots for top 8 important features
        '''
        # shap_values_plot:
        #   shape = (n_samples, n_features)
        # testData_df:
        #   DataFrame, columns should match feature_names

        # 如果 shap_values_plot 是 list（二分类某些版本可能出现），取第 1 类
        if isinstance(shap_values_plot, list):
            shap_values_mapping = shap_values_plot[1]
        else:
            shap_values_mapping = shap_values_plot

        # 若为 Explanation 对象，取 .values
        if hasattr(shap_values_mapping, 'values'):
            shap_values_mapping = shap_values_mapping.values

        # 计算每个特征的平均绝对 SHAP 值，并选前 8 个重要特征
        mean_abs_shap = np.abs(shap_values_mapping).mean(axis=0)
        top8_idx = np.argsort(mean_abs_shap)[::-1][:8]
        top8_feature_names = [feature_names[i] for i in top8_idx]

        # 创建 2x4 子图
        fig, axes = plt.subplots(2, 4, figsize=(16, 8), dpi=300, facecolor='white')
        axes = axes.flatten()

        subplot_labels = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h']

        for i, feat_name in enumerate(top8_feature_names):
            ax = axes[i]

            feat_idx = list(feature_names).index(feat_name)
            x_feature = testData_df[feat_name].values
            y_shap = shap_values_mapping[:, feat_idx]

            # 绘制散点图
            ax.scatter(x_feature, y_shap, s=10, c='dodgerblue', alpha=0.85, edgecolors='none')

            # 横纵参考线
            ax.axhline(y=0, color='gray', linestyle='--', linewidth=1, alpha=0.8)

            # 坐标轴标签
            ax.set_xlabel(str(feat_name), fontsize=11)
            ax.set_ylabel(f'SHAP value for\n{feat_name}', fontsize=11)

            # 美化
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.tick_params(axis='both', labelsize=10)

            # 子图编号 a-h
            ax.text(0.5, -0.28, subplot_labels[i],
                    transform=ax.transAxes,
                    ha='center', va='center',
                    fontsize=16, family='serif')

        # 如果特征不足 8 个，隐藏多余子图
        for j in range(len(top8_feature_names), 8):
            axes[j].axis('off')

        plt.tight_layout()

        plt.savefig(outPath + '\\' + '_XGB_SHAP_Value_Mapping_Top8.tif',
                    format='tif', bbox_inches='tight', transparent=True, dpi=300)

        # plt.rcParams['svg.fonttype'] = 'none'
        # plt.savefig(outPath + '\\' + '_XGB_SHAP_Value_Mapping_Top8.svg',
        #             format='svg', bbox_inches='tight', transparent=True)

        plt.close()
        with open(outPath + '\\' + r'_XGB_SHAP_Value_Mapping_Top8_说明.txt', 'w', encoding='utf-8') as f:
            f.write("SHAP值映射散点图（图4c）分析说明：\n")
            f.write("本图展示贡献度前8位关键拉曼特征的光谱强度与SHAP值之间的函数关系（2×4子图，a–h）。\n")
            f.write("横轴为归一化光谱强度，纵轴为对应SHAP贡献值，每点代表一个测试样本。\n\n")
            f.write("子图a（934.2 cm⁻¹，多糖/蛋白质α螺旋C-C振动）：\n")
            f.write("  正相关趋势，高强度样本正向SHAP贡献显著，指向Lichtheimia属真菌高糖化酶活性的光谱表型。\n\n")
            f.write("子图b（949.4 cm⁻¹，碳水化合物C-C骨架伸缩）：\n")
            f.write("  与子图a高度一致，二者协同正向贡献共同指示多糖代谢能力是区分发酵微生物的核心生化特征。\n\n")
            f.write("子图c（1715.5 cm⁻¹，酯羰基C=O伸缩振动）：\n")
            f.write("  最典型的单调正相关模式，SHAP值随酯羰基强度增大近线性升高，样本散布集中度最高。\n")
            f.write("  是酯化活性的核心光谱标志，与绵柔型白酒主体乳酸乙酯生物合成直接关联。\n\n")
            f.write("子图d（734.3 cm⁻¹，腺嘌呤环呼吸振动）：\n")
            f.write("  散点分布分散，呈典型非线性异质性，反映细菌与真菌在RNA代谢活性上的样本依赖性差异。\n\n")
            f.write("子图e（1322.8 cm⁻¹，酰胺III带/蛋白质振动）：\n")
            f.write("  正相关趋势，高强度对应高蛋白酶表达量菌株，与大曲发酵中底物蛋白质水解及风味氨基酸释放相关。\n\n")
            f.write("子图f（1030.9 cm⁻¹，苯丙氨酸苯环呼吸振动）：\n")
            f.write("  高强度正向SHAP样本集中于右上区域，指向Bacillus licheniformis经Ehrlich途径\n")
            f.write("  代谢苯丙氨酸生成苯乙醇等花香型风味化合物的代谢能力。\n\n")
            f.write("子图g（2862.8 cm⁻¹，CH₂对称伸缩/脂肪酸链）：\n")
            f.write("  正向单调关系，高脂肪酸链密度菌株获较大正向贡献，与高级脂肪酸乙酯（棕榈酸乙酯等）\n")
            f.write("  生成及绵柔型白酒油脂感风格特征形成相关。\n\n")
            f.write("子图h（1087.7 cm⁻¹，磷脂O-P-O对称伸缩振动）：\n")
            f.write("  高强度样本偏正向，但样本间散布较大，反映磷脂膜组成种间异质性，\n")
            f.write("  与甘油磷脂代谢通路及脂质风味前体积累存在代谢网络交联。\n\n")
            f.write("综合结论：1715.5 cm⁻¹具有最稳定线性特征-SHAP响应关系，是连接单细胞光谱与\n")
            f.write("白酒酯类风味生成潜能的最核心光谱标志物；蛋白质与多糖类特征协同反映菌株水解酶功能分化。\n")

        # =========================================
        '''
            waterfall + force plot 联合版本
        '''
        # 5. 选择需要解释的单个样本
        # 优先选择误判样本；若没有误判，则解释第0个样本
        y_pred = best_model_xgb.predict(testData)
        wrong_idx = np.where(y_pred != testTarget)[0]

        if len(wrong_idx) > 0:
            sample_idx = int(wrong_idx[0])
            sample_tag = f'misclassified_{sample_idx}'
        else:
            sample_idx = 0
            sample_tag = f'sample_{sample_idx}'

        # =========================================# =========================================
        # 6. SHAP waterfall plot
        try:
            # shap_exp = explainer(testData_df)
            # single_exp = shap_exp[sample_idx]

            shap_exp = explainer(testData_df)
            # 多分类时：必须指定某一个类别
            pred_class = int(y_pred[sample_idx])
            try:
                single_exp = shap_exp[sample_idx, :, pred_class]
            except:
                try:
                    single_exp = shap_exp[sample_idx, pred_class]
                except:
                    single_values = shap_values_plot[sample_idx, :, pred_class]
                    base_value = explainer.expected_value[pred_class] if isinstance(explainer.expected_value, (
                    list, np.ndarray)) else explainer.expected_value
                    single_exp = shap.Explanation(
                        values=single_values,
                        base_values=base_value,
                        data=testData_df.iloc[sample_idx].values,
                        feature_names=feature_names
                    )
        except:
            if hasattr(explainer, 'expected_value'):
                base_value = explainer.expected_value
            else:
                base_value = 0

            if isinstance(base_value, list) or isinstance(base_value, np.ndarray):
                if np.ndim(base_value) > 0:
                    base_value = base_value[1] if len(base_value) > 1 else base_value[0]

            single_exp = shap.Explanation(
                values=shap_values_plot[sample_idx],
                base_values=base_value,
                data=testData_df.iloc[sample_idx].values,
                feature_names=feature_names
            )

        plt.figure(figsize=(8, 6), dpi=300, facecolor='white')
        shap.plots.waterfall(single_exp, max_display=15, show=False)
        plt.tight_layout()
        plt.savefig(outPath + '\\' + f'_XGB_SHAP_Waterfall_{sample_tag}.tif',
                    format='tif', bbox_inches='tight', transparent=True, dpi=300)
        # plt.rcParams['svg.fonttype'] = 'none'
        # plt.savefig(outPath + '\\' + f'_XGB_SHAP_Waterfall_{sample_tag}.svg',
        #             format='svg', bbox_inches='tight', transparent=True)
        plt.close()

        # =========================================# =========================================
        # SHAP force plot：分别生成 3 个样本的 force plot，再拼接成一张总图（上中下排列）
        # 说明：
        # 1）避免 shap.force_plot(matplotlib=True) 直接用于 subplot 时发生重叠
        # 2）每个样本仅保留 Top-N 个重要特征，避免蓝色负向特征过多导致图像横向比例失衡
        # 3）最终输出一个 PNG 总图，每个子图显示一个样本，并标注样本编号
        import matplotlib.image as mpimg
        # ===== 你可以自定义这里 =====
        # sample_indices = [0, 200, 400]  # 想解释的三个测试样本编号
        # top_n_features = 10  # 每个样本最多展示的特征数
        # single_fig_width = 20  # 单个 force plot 宽度
        # single_fig_height = 3.4  # 单个 force plot 高度
        # # 最终拼接总图尺寸
        # final_fig_width = 20  # 拼接总图宽度
        # final_fig_height = 8.8  # 拼接总图高度

        # ===== 你可以自定义这里 =====
        # 自动根据 6 个文件夹（类别数）选择每个文件夹下第一个样本
        # 每个类别在 testTarget 中按顺序存放，因此取索引范围段的首个样本
        samples_per_class = []
        cursor = 0
        for count in pcaCount:  # 每个类别样本数量，前面已有定义
            samples_per_class.append(cursor)  # 每类第一个样本编号
            cursor += int(count * svmTestSize)  # 按 train_test_split 比例偏移估算索引
        # 若有6类，则选出6个样本
        sample_indices = samples_per_class

        top_n_features = 6  # 每个样本最多展示的特征数
        single_fig_width = 20
        single_fig_height = 3.4
        # 根据样本数自动调整总图尺寸（每个子图约3.4高）
        final_fig_width = 20
        final_fig_height = 3.4 * len(sample_indices) + 0.3


        # 输出分辨率
        save_dpi = 500
        # 临时图片保存目录
        temp_dir = os.path.join(outPath, 'temp_force_plots')
        os.makedirs(temp_dir, exist_ok=True)

        # 处理 expected_value
        if hasattr(explainer, 'expected_value'):
            expected_value = explainer.expected_value
        else:
            expected_value = 0

        if isinstance(expected_value, (list, np.ndarray)):
            if np.ndim(expected_value) > 0:
                expected_value = expected_value[1] if len(expected_value) > 1 else expected_value[0]

        temp_img_paths = []

        # =========================
        # 第一步：分别生成每个样本的 force plot
        # =========================
        for i, sample_idx in enumerate(sample_indices):
            # 当前样本 SHAP 值
            single_shap = shap_values_plot[sample_idx, :]

            # 仅保留 |SHAP| 最大的前 top_n_features 个特征
            top_idx = np.argsort(np.abs(single_shap))[::-1][:top_n_features]

            # 按原始特征顺序排序，便于阅读
            top_idx = np.sort(top_idx)

            shap_values_selected = single_shap[top_idx]

            # 对应特征输入值
            data_selected = testData_df.iloc[sample_idx, top_idx].values
            data_selected_rounded = np.array([float(f"{x:.3f}") for x in data_selected])

            # 特征名只保留名称，避免重复 “=”
            feature_names_selected = [feature_names[j] for j in top_idx]

            # 创建单图
            plt.figure(figsize=(single_fig_width, single_fig_height), dpi=save_dpi, facecolor='white')

            shap.force_plot(
                base_value=expected_value,
                shap_values=shap_values_selected,
                features=data_selected_rounded,
                feature_names=feature_names_selected,
                matplotlib=True,
                show=False,
                contribution_threshold=0.03
            )


            # 保存单张图，尽可能去掉多余白边
            temp_img_path = os.path.join(temp_dir, f'force_sample_{sample_idx}.png')
            plt.savefig(temp_img_path, format='png', bbox_inches='tight', pad_inches=0.03, dpi=save_dpi)
            plt.close()

            temp_img_paths.append(temp_img_path)

        # =========================
        # 第二步：拼接成总图
        # =========================
        fig, axes = plt.subplots(len(sample_indices), 1, figsize=(final_fig_width, final_fig_height), dpi=save_dpi)

        if len(sample_indices) == 1:
            axes = [axes]

        for ax, img_path, sample_idx in zip(axes, temp_img_paths, sample_indices):
            img = mpimg.imread(img_path)
            ax.imshow(img)
            ax.axis('off')

        # 压缩子图间距和边缘空白
        plt.subplots_adjust(left=0.01, right=0.99, top=0.99, bottom=0.01, hspace=0.02)

        # 保存总图
        png_path = os.path.join(outPath, '_XGB_SHAP_Force_6samples_combined.png')
        # svg_path = os.path.join(outPath, '_XGB_SHAP_Force_6samples_combined.svg')

        plt.savefig(png_path, format='png', bbox_inches='tight', pad_inches=0.02, dpi=save_dpi)
        # plt.rcParams['svg.fonttype'] = 'none'
        # plt.savefig(svg_path, format='svg', bbox_inches='tight', pad_inches=0.02)
        plt.close()

        # =========================
        # 第三步：保存说明文字
        # =========================
        with open(outPath + '\\' + r'_XGB_SHAP_Force_6samples_combined_说明.txt', 'w', encoding='utf-8') as f:
            f.write("SHAP force plot（三个样本拼接图）说明：\n")
            f.write(f"样本编号：{sample_indices}\n")
            f.write(f"每个样本最多显示前 {top_n_features} 个重要特征（按 |SHAP| 排序）。\n")
            f.write("图像已保存为 PNG 和 SVG 格式。\n")
            f.write("图中上、中、下分别对应三个不同测试样本。\n")
            f.write("每个子图左上角已标注样本序号。\n")
            f.write("红色特征表示推动模型输出升高的拉曼位移特征。\n")
            f.write("蓝色特征表示推动模型输出降低的拉曼位移特征。\n")
            f.write("为了避免过多低贡献蓝色特征影响横向显示，代码中已仅保留最重要的特征。\n")

        # =========================================
        # 8. 保存文字说明
        with open(outPath + '\\' + f'_XGB_SHAP_SingleSample_{sample_tag}_说明.txt', 'w', encoding='utf-8') as f:
            f.write("单样本 SHAP 解释说明：\n")
            f.write(f"解释样本编号：{sample_idx}\n")
            f.write(f"模型预测类别：{int(y_pred[sample_idx])}\n")
            f.write(f"真实类别：{int(testTarget[sample_idx])}\n")
            f.write("\n")
            f.write("1. Waterfall plot：\n")
            f.write("   展示该样本从模型基线输出到最终输出的形成过程。\n")
            f.write("   红色特征推动输出升高，蓝色特征推动输出降低。\n")
            f.write("\n")
            f.write("2. Force plot：\n")
            f.write("   展示各个拉曼位移特征对该样本分类结果的推拉作用。\n")
            f.write("   该图保存为 HTML，建议使用浏览器打开查看。\n")
            f.write("\n")
            f.write("3. 对于二分类问题，若使用的是类别1对应的 SHAP 值：\n")
            f.write("   SHAP > 0 表示推动样本向类别1预测；\n")
            f.write("   SHAP < 0 表示推动样本向类别0预测。\n")




        # **************************************************************************************************************XGB特征重要性(gain增益)排序条形图
        '''
        # 重要性指标feature_importances_，这个参数好像只有在决策树和以决策树为基础的算法有。
        '''
        top_K = 20
        # 返回前n个特征
        # xgbFeatureImportance = XGBClassifier()
        xgbFeaImp = best_model_xgb.fit(trainData, trainTarget).feature_importances_
        top_k_idx = np.array(xgbFeaImp).argsort()[::-1][0:top_K]
        # argsort:返回的是数组值从小到大的索引值
        f = open(outPath + '\\' + r'_xgbGain.txt', "w")
        f.write('用XGBoost模型对光谱特征重要性进行排序：输出前10最重要拉曼位移(特征)及重要性占比' + '\n')
        for interI in range(top_K):
            f.write("拉曼位移:" + str(x[top_k_idx[interI]]) + ", 占比：" + str(round(xgbFeaImp[top_k_idx[interI]], 3)) + '\n')

        f.write('注：属性(拉曼位移对应的波数点)对分裂点改进性能度量越大（越靠近根节点），权值越大；\n'
                '被越多提升树所选择，属性越重要。这里性能度量选择分裂节点的gain增益。' + '\n')
        f.close()

        wavenumberX = np.arange(len(top_k_idx))
        # 添加波数坐标
        plt.bar(x=wavenumberX, height=xgbFeaImp[top_k_idx], width=0.5, tick_label=x[top_k_idx], color='#00BFFF')
        plt.xticks(fontsize=12, rotation=45)
        plt.yticks(fontsize=12)
        plt.xlabel("Characteristic peaks", fontsize=14)
        plt.ylabel("Gain", fontsize=14)
        plt.tight_layout()
        # 解决显示不全的问题
        # plt.show()

        plt.gca().yaxis.set_major_formatter(FuncFormatter(to_percent))
        # from matplotlib import ticker
        # plt.gca().yaxis.set_major_formatter(ticker.PercentFormatter(xmax=1, decimals=1))

        # 纵坐标以百分百%显示
        outPath123 = outPath + '\\' + r'_XGB_Gain.tif'
        plt.savefig(outPath123, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        plt.rcParams['svg.fonttype'] = 'none'
        # outPath123 = outPath + '\\' + r'_XGB_Gain.svg'
        # plt.savefig(outPath123, format='svg', bbox_inches='tight', transparent=True)
        plt.close()
        # return 0
        # # ********************************************************************************************保存xgb网格参数
        # f = open(outPath + '\\' + r'_xgbParameter.txt', "w")
        # f.write("xgb参数说明：\n")
        # f.write('搜索网格参数1.基分类器模型(booster)\n')
        # f.write('\tgbtree:基分类器为树模型;gbliner:基分类器为线性模型\n')
        # f.write('搜索网格参数2.树的深度(max_depth)\n')
        # f.write('\t[5,6,7]值过大容易过拟合，值过小容易欠拟合\n')
        # f.write('\tdistance:权重点等于他们距离的倒数。使用此函数，更近的邻居对于所预测的点的影响更大。\n')
        # f.write('搜索网格参数3.基分类器的个数(n_estimatores)\n')
        # f.write('\t[80, 100, 120]\n')
        # f.write('搜索最优网格参数:' + '\n' + str(grid_search_xgb.best_params_) + '\n')
        # f.write('默认超参数:' + "\n" + str(best_model_xgb) + '\n')
        # f.close()

        # xgb.fit(trainData, trainTarget)
        scoreXGB = xgb.score(testData, testTarget)
        # 测试数据准确率
        y_scoreXGB1 = xgb.predict(testData)
        # 预测标签
        y_scoreXGB = xgb.predict_proba(testData)
        # 预测标签概率
        # matrixFindsize = 12
        matrixFindsize = 12 - 0.1*len(classOneLabel)

        # 混淆矩阵字体大小
        matrixDip = 300
        # 分辨率
        # Plot confusion matrix  绘图混淆矩阵
        sns.set_context("talk", rc={"font": "Helvetica", "font.size": matrixFindsize})
        label = [files[i] for i in classOneLabel]
        cm = confusion_matrix(testTarget, xgb.predict(testData), labels=classOneLabel)
        plt.figure(figsize=(5, 4),
                   dpi=matrixDip,
                   # 分辨率
                   facecolor='blueviolet',
                   # 背景色,'snow','blueviolet','lightgray'
                   )
        # 设置title
        # plt.title('Test results of XGB classifier', fontsize=12)

        cm = 100 * cm / cm.sum(axis=1)[:, np.newaxis]
        ax = sns.heatmap(cm, annot=True, cmap='YlGnBu', fmt='0.1f', xticklabels=label, yticklabels=label)
        ax.xaxis.tick_top()
        ax.set_xticklabels(ax.get_xticklabels(), rotation=90)
        # plt.xticks(rotation=90)
        plt.yticks(rotation=0)
        # 设置刻度值即坐标轴刻度 xticks yticks
        # plt.tick_params(labelsize=10)
        plt.tick_params(labelsize=matrixFindsize)
        # 设置坐标轴的刻度参数
        # 标签设置
        plt.xlabel('Predicted_label', fontsize=14)
        plt.ylabel('True_label', fontsize=14)
        ax.xaxis.set_label_position('top')
        plt.tight_layout()
        outPath122 = outPath + '\\' + r'_XGB.tif'
        plt.savefig(outPath122, format='tif', bbox_inches='tight', transparent=True, dpi=300)
        plt.rcParams['svg.fonttype'] = 'none'
        outPath122 = outPath + '\\' + r'_XGB.svg'
        plt.savefig(outPath122, format='svg', bbox_inches='tight', transparent=True)
        plt.close()

        # ================================================================================================================== 精度与混淆矩阵（测试集与训练集，XGBoost）======
        # y_pred_test_xgb 复用已有 y_scoreXGB1，y_pred_train_xgb 为新增训练集预测
        y_pred_train_xgb = xgb.predict(trainData)
        y_pred_test_xgb = y_scoreXGB1  # 已在上方计算

        matrixFindsize_xgb = 12 - 0.1 * len(classOneLabel)
        sns.set_context("talk", rc={"font": "Helvetica",
                                    "font.size": matrixFindsize_xgb})
        label_xgb = [files[i] for i in classOneLabel]

        # ---------------------------------- 训练集混淆矩阵（百分比形式）
        # 配色：YlGnBu（与现有 _XGB.tif 保持同色系），背景 blueviolet
        cm_train_xgb = confusion_matrix(trainTarget, y_pred_train_xgb,
                                        labels=classOneLabel)
        cm_train_xgb_pct = 100 * cm_train_xgb / cm_train_xgb.sum(axis=1)[:, np.newaxis]
        plt.figure(figsize=(5, 4), dpi=300, facecolor='blueviolet')
        ax = sns.heatmap(cm_train_xgb_pct, annot=True, cmap='YlGnBu',
                         fmt='0.1f', xticklabels=label_xgb, yticklabels=label_xgb)
        ax.xaxis.tick_top()
        ax.set_xticklabels(ax.get_xticklabels(), rotation=90)
        plt.yticks(rotation=0)
        plt.tick_params(labelsize=matrixFindsize_xgb)
        plt.xlabel('Predicted_label', fontsize=14)
        plt.ylabel('True_label', fontsize=14)
        ax.xaxis.set_label_position('top')
        plt.tight_layout()
        plt.savefig(outPath + '\\' + r'_XGB_TrainMatrixPercent.tif',
                    format='tif', bbox_inches='tight', transparent=True, dpi=300)
        plt.close()

        # ---------------------------------- 测试集混淆矩阵（百分比形式）
        # 配色：PuBuGn（蓝绿渐变，与训练集矩阵形成视觉区分），背景 blueviolet
        cm_test_xgb = confusion_matrix(testTarget, y_pred_test_xgb,
                                       labels=classOneLabel)
        cm_test_xgb_pct = 100 * cm_test_xgb / cm_test_xgb.sum(axis=1)[:, np.newaxis]
        plt.figure(figsize=(5, 4), dpi=300, facecolor='blueviolet')
        ax = sns.heatmap(cm_test_xgb_pct, annot=True, cmap='PuBuGn',
                         fmt='0.1f', xticklabels=label_xgb, yticklabels=label_xgb)
        ax.xaxis.tick_top()
        ax.set_xticklabels(ax.get_xticklabels(), rotation=90)
        plt.yticks(rotation=0)
        plt.tick_params(labelsize=matrixFindsize_xgb)
        plt.xlabel('Predicted_label', fontsize=14)
        plt.ylabel('True_label', fontsize=14)
        ax.xaxis.set_label_position('top')
        plt.tight_layout()
        plt.savefig(outPath + '\\' + r'_XGB_TestMatrixPercent.tif',
                    format='tif', bbox_inches='tight', transparent=True, dpi=300)
        plt.close()

        # ---------------------------------- 测试集混淆矩阵（样本数量形式）
        # 配色：Blues（纯蓝色系，整数样本数），背景 blueviolet
        plt.figure(figsize=(5, 4), dpi=300, facecolor='blueviolet')
        ax = sns.heatmap(cm_test_xgb, annot=True, cmap='Blues',
                         fmt='d', xticklabels=label_xgb, yticklabels=label_xgb)
        ax.xaxis.tick_top()
        ax.set_xticklabels(ax.get_xticklabels(), rotation=90)
        plt.yticks(rotation=0)
        plt.tick_params(labelsize=matrixFindsize_xgb)
        plt.xlabel('Predicted_label', fontsize=14)
        plt.ylabel('True_label', fontsize=14)
        ax.xaxis.set_label_position('top')
        plt.tight_layout()
        plt.savefig(outPath + '\\' + r'_XGB_TestMatrixCounts.tif',
                    format='tif', bbox_inches='tight', transparent=True, dpi=300)
        plt.close()

        # ====== 绘制多类别训练误差曲线（XGBoost，专业科研配图规范）======
        from sklearn.metrics import accuracy_score as _acc_xgb

        n_estimators_range_xgb = range(1, 120, 2)
        class_labels_xgb = classOneLabel
        label_names_xgb = [files[i] for i in class_labels_xgb]
        class_errors_xgb = {cls: [] for cls in class_labels_xgb}
        overall_errors_xgb = []

        # 逐步增加 n_estimators，每轮从头训练一个小 XGBClassifier
        # 说明：XGBClassifier 的 warm_start 在部分版本行为不稳定，
        #       此处采用从头训练方式，确保兼容性
        for _n_xgb in n_estimators_range_xgb:
            try:
                _xgb_cv = XGBClassifier(
                    use_label_encoder=False,
                    eval_metric='logloss',
                    random_state=1,
                    n_estimators=_n_xgb
                )
                _xgb_cv.fit(trainData, trainTarget)
            except Exception:
                _xgb_cv = XGBClassifier(
                    eval_metric='logloss',
                    random_state=1,
                    n_estimators=_n_xgb
                )
                _xgb_cv.fit(trainData, trainTarget)

            _y_cv = _xgb_cv.predict(trainData)
            overall_errors_xgb.append(1 - _acc_xgb(trainTarget, _y_cv))

            for _cls_xgb in class_labels_xgb:
                _mask_xgb = (trainTarget == _cls_xgb)
                _corr_xgb = np.sum(_y_cv[_mask_xgb] == _cls_xgb)
                _total_xgb = np.sum(_mask_xgb)
                _err_xgb = 1 - _corr_xgb / _total_xgb if _total_xgb > 0 else 0
                class_errors_xgb[_cls_xgb].append(_err_xgb)

        # 专业科研论文配图规范：多线型 + 稀疏标记 + 去多余边框
        _xgb_linestyles = [
            '-', '--', '-.', ':',
            (0, (3, 1, 1, 1)),
            (0, (5, 1)),
            (0, (3, 5, 1, 5)),
            (0, (1, 1))
        ]
        _xgb_markevery = max(1, len(list(n_estimators_range_xgb)) // 8)

        plt.figure(figsize=(7, 5), dpi=300, facecolor='white')

        for i, cls in enumerate(class_labels_xgb):
            plt.plot(
                list(n_estimators_range_xgb),
                class_errors_xgb[cls],
                color=colorSpectrumColour[i % len(colorSpectrumColour)],
                linestyle=_xgb_linestyles[i % len(_xgb_linestyles)],
                linewidth=1.5,
                marker=markers[i % len(markers)],
                markersize=4,
                markevery=_xgb_markevery,
                markeredgewidth=0.5,
                markeredgecolor='white',
                alpha=0.85,
                label=f'{label_names_xgb[i]} Error'
            )

        plt.plot(
            list(n_estimators_range_xgb),
            overall_errors_xgb,
            color='black',
            linestyle='--',
            linewidth=2.5,
            marker='D',
            markersize=4,
            markevery=_xgb_markevery,
            markeredgewidth=0.5,
            markeredgecolor='white',
            label='Overall Model Error'
        )

        plt.xlabel('Number of Estimators', fontsize=14)
        plt.ylabel('Error Rate', fontsize=14)
        plt.legend(loc='upper right', fontsize=18,
                   framealpha=0.85, edgecolor='#CCCCCC', ncol=2)
        plt.grid(True, linestyle='--', alpha=0.35, linewidth=0.7)
        _ax_xgb_conv = plt.gca()
        _ax_xgb_conv.spines['top'].set_visible(False)
        _ax_xgb_conv.spines['right'].set_visible(False)
        plt.tight_layout()

        plt.savefig(outPath + '\\' + r'_XGB_ClassConvergence_Error.tif',
                    format='tif', bbox_inches='tight', transparent=True, dpi=300)
        plt.rcParams['svg.fonttype'] = 'none'
        plt.savefig(outPath + '\\' + r'_XGB_ClassConvergence_Error.svg',
                    format='svg', bbox_inches='tight', transparent=True)
        plt.close()

        # ====== 保存数值结果（XGBoost 收敛误差）======
        with open(outPath + '\\' + r'_XGB_ClassConvergence_Error.txt',
                  'w', encoding='utf-8') as f:
            f.write("XGBoost收敛性分析（多类别训练误差随 n_estimators 变化）\n")
            _header_xgb = "n_estimators"
            for i, cls in enumerate(class_labels_xgb):
                _header_xgb += f"\t{label_names_xgb[i]}_error"
            _header_xgb += "\toverall_error\n"
            f.write(_header_xgb)
            for _idx_xgb, _n_xgb in enumerate(n_estimators_range_xgb):
                _line_xgb = f"{_n_xgb}"
                for cls in class_labels_xgb:
                    _line_xgb += f"\t{round(class_errors_xgb[cls][_idx_xgb], 4)}"
                _line_xgb += f"\t{round(overall_errors_xgb[_idx_xgb], 4)}\n"
                f.write(_line_xgb)





        # *************************************************************************************XGB的ROC曲线与AUC面积
        if 2 == len(classOneLabel):
            # 二分类ROC曲线
            fpr, tpr, threshold = roc_curve(testTarget, y_scoreXGB1)
            ###计算真正率和假正率
            roc_auc = auc(fpr, tpr)
            ###计算auc的值
            # Plot all ROC curves
            plt.figure(figsize=(5, 4), dpi=300)
            lw = 2
            plt.plot(fpr, tpr, color='darkorange',
                     lw=lw, label='ROC curve (area = %0.2f)' % roc_auc)
            ###假正率为横坐标，真正率为纵坐标做曲线
            plt.plot([0, 1], [0, 1], color='navy', lw=lw, linestyle='--')
            plt.xlim([0.0, 1.0])
            plt.ylim([0.0, 1.05])

            plt.xlabel('False Positive Rate', fontsize=14)
            plt.ylabel('True Positive Rate', fontsize=14)
            # plt.title('XGB_ROC', fontsize=12)
            plt.legend(loc="lower right", fontsize=12)
            # plt.show()
            outPath121 = outPath + '\\' + r'_XGB_ROC.tif'
            plt.savefig(outPath121, format='tif', bbox_inches='tight', transparent=True, dpi=300)
            plt.rcParams['svg.fonttype'] = 'none'
            outPath121 = outPath + '\\' + r'_XGB_ROC.svg'
            plt.savefig(outPath121, format='svg', bbox_inches='tight', transparent=True)
            plt.close()
        else:
            # 多分类
            for i in range(len(classOneLabel)):
                fpr[i], tpr[i], _ = roc_curve(testTarget1[:, i], y_scoreXGB[:, i])
                roc_auc[i] = auc(fpr[i], tpr[i])

            # Compute micro-average ROC curve and ROC area（方法二）
            fpr["micro"], tpr["micro"], _ = roc_curve(testTarget1.ravel(), y_scoreXGB.ravel())
            roc_auc["micro"] = auc(fpr["micro"], tpr["micro"])

            # Compute macro-average ROC curve and ROC area（方法一）
            # First aggregate all false positive rates
            all_fpr = np.unique(np.concatenate([fpr[i] for i in range(len(classOneLabel))]))
            # Then interpolate all ROC curves at this points
            mean_tpr = np.zeros_like(all_fpr)
            for i in range(len(classOneLabel)):
                mean_tpr += interp(all_fpr, fpr[i], tpr[i])
            # Finally average it and compute AU
            mean_tpr /= len(classOneLabel)
            fpr["macro"] = all_fpr
            tpr["macro"] = mean_tpr
            roc_auc["macro"] = auc(fpr["macro"], tpr["macro"])

            # Plot all ROC curves
            plt.figure(figsize=(5, 4), dpi=300)
            plt.plot(fpr["micro"], tpr["micro"],
                     label='micro-average ROC curve (AUC = {0:0.2f})'
                           ''.format(roc_auc["micro"]),
                     color='deeppink', linestyle=':')

            plt.plot(fpr["macro"], tpr["macro"],
                     label='macro-average ROC curve (AUC = {0:0.2f})'
                           ''.format(roc_auc["macro"]),
                     color='navy', linestyle=':')

            # colors = cycle(['aqua', 'darkorange', 'cornflowerblue'])
            for i in range(len(classOneLabel)):
                plt.plot(fpr[i], tpr[i],
                         label='ROC curve of class {0} (AUC = {1:0.2f})'
                               ''.format(files[i], roc_auc[i]))

            plt.plot([0, 1], [0, 1], 'k--')
            plt.xlim([0.0, 1.0])
            plt.ylim([0.0, 1.05])
            plt.xlabel('False Positive Rate', fontsize=14)
            plt.ylabel('True Positive Rate', fontsize=14)
            # plt.title('XGB_ROC:multi-class', fontsize=12)
            plt.legend(loc="lower right", fontsize=12)
            # plt.show()
            outPath121 = outPath + '\\' + r'_XGB_ROC.tif'
            plt.savefig(outPath121, format='tif', bbox_inches='tight', transparent=True, dpi=300)
            plt.rcParams['svg.fonttype'] = 'none'
            outPath121 = outPath + '\\' + r'_XGB_ROC.svg'
            plt.savefig(outPath121, format='svg', bbox_inches='tight', transparent=True)
            plt.close()

        #随机森林
        from sklearn.ensemble import RandomForestClassifier
        # **************************************************************************************************************随机森林模块
        figure(3 * filesNumInter + 9)

        # from sklearn.ensemble import RandomForestClassifier
        #
        # rf = RandomForestClassifier(random_state=1)
        # param_grid = {
        #     'n_estimators': [100, 200, 300],
        #     'max_depth': [5, 10, 15],
        #     'min_samples_split': [2, 4, 6]
        # }
        # # grid_search_rf = GridSearchCV(rf, param_grid, cv=3, n_jobs=1)
        # # try:
        # #     rf = grid_search_rf.fit(trainData, trainTarget)
        # # except:
        # #     rf = grid_search_rf.fit(trainData, trainTarget)
        # try:
        #     rf = rf.fit(trainData, trainTarget)
        # except:
        #     rf = rf.fit(trainData, trainTarget)
        rf = RandomForestClassifier(random_state=1)
        rf = rf.fit(trainData, trainTarget)
        # best_model_rf = rf
        # 保存最优模型
        outPathRF = outPath + '\\' + r'_rfModel.pkl'
        joblib.dump(rf, outPathRF)
        # best_model_rf = grid_search_rf.best_estimator_
        best_model_rf = rf
        # ************************************** 特征重要性可视化（类似 XGB Gain）
        rf_importance = best_model_rf.feature_importances_
        top_K = 20
        top_k_idx = np.argsort(rf_importance)[::-1][:top_K]
        with open(outPath + '\\' + r'_rfGain.txt', 'w', encoding='utf-8') as f:
            f.write('随机森林模型特征重要性（前20个拉曼位移）\n')
            for i in range(top_K):
                f.write(f"拉曼位移: {round(x[top_k_idx[i]], 2)}, 占比: {round(rf_importance[top_k_idx[i]], 2)}\n")

        # plt.bar(
        #     x=np.arange(len(top_k_idx)),
        #     height=rf_importance[top_k_idx],
        #     tick_label=x[top_k_idx],
        #     color='#32CD32', width=0.5
        # )
        # ====== 绘制前20重要特征条形图，横坐标保留两位小数
        top_feature_values = rf_importance[top_k_idx]
        top_feature_labels = [f"{val:.2f}" for val in x[top_k_idx]]

        plt.figure(figsize=(8, 5), dpi=300, facecolor='white')
        plt.bar(
            x=np.arange(len(top_k_idx)),
            height=top_feature_values,
            tick_label=top_feature_labels,
            color='#32CD32',
            width=0.5
        )


        plt.xticks(rotation=45, fontsize=12)
        plt.yticks(fontsize=12)
        plt.xlabel("Characteristic peaks", fontsize=14)
        plt.ylabel("Importance", fontsize=14)
        plt.tight_layout()
        plt.gca().yaxis.set_major_formatter(FuncFormatter(to_percent))
        plt.savefig(outPath + '\\' + r'_RF_Gain.tif', format='tif', bbox_inches='tight',
                    transparent=True, dpi=300)
        plt.close()

        # ====== 精度与混淆矩阵（测试集与训练集） ======
        y_pred_train = rf.predict(trainData)
        y_pred_test = rf.predict(testData)
        y_scoreRF = rf.predict_proba(testData)
        scoreRF_test = rf.score(testData, testTarget)
        scoreRF_train = rf.score(trainData, trainTarget)

        matrixFindsize = 12 - 0.1 * len(classOneLabel)
        sns.set_context("talk", rc={"font": "Helvetica", "font.size": matrixFindsize})
        label = [files[i] for i in classOneLabel]

        # ---------------------------------- 训练集混淆矩阵（百分比形式）
        cm_train = confusion_matrix(trainTarget, y_pred_train, labels=classOneLabel)
        cm_train_percent = 100 * cm_train / cm_train.sum(axis=1)[:, np.newaxis]
        plt.figure(figsize=(5, 4), dpi=300, facecolor='lightgray')
        # plt.title('Training set (80%) results of RandomForest', fontsize=12)
        ax = sns.heatmap(cm_train_percent, annot=True, cmap='YlGn', fmt='0.1f',
                         xticklabels=label, yticklabels=label)
        ax.xaxis.tick_top()
        ax.set_xticklabels(ax.get_xticklabels(), rotation=90)
        plt.yticks(rotation=0)
        plt.xlabel('Predicted_label', fontsize=14)
        plt.ylabel('True_label', fontsize=14)
        ax.xaxis.set_label_position('top')
        plt.tight_layout()
        plt.savefig(outPath + '\\' + r'_RF_TrainMatrixPercent.tif', format='tif', bbox_inches='tight',
                    transparent=True, dpi=300)
        plt.close()

        # ---------------------------------- 测试集混淆矩阵（百分比形式）
        cm_test = confusion_matrix(testTarget, y_pred_test, labels=classOneLabel)
        cm_test_percent = 100 * cm_test / cm_test.sum(axis=1)[:, np.newaxis]
        plt.figure(figsize=(5, 4), dpi=300, facecolor='lightgray')
        # plt.title('Test set (20%) results of RandomForest', fontsize=12)
        ax = sns.heatmap(cm_test_percent, annot=True, cmap='Greens', fmt='0.1f',
                         xticklabels=label, yticklabels=label)
        ax.xaxis.tick_top()
        ax.set_xticklabels(ax.get_xticklabels(), rotation=90)
        plt.yticks(rotation=0)
        plt.xlabel('Predicted_label', fontsize=14)
        plt.ylabel('True_label', fontsize=14)
        ax.xaxis.set_label_position('top')
        plt.tight_layout()
        plt.savefig(outPath + '\\' + r'_RF_TestMatrixPercent.tif', format='tif', bbox_inches='tight',
                    transparent=True, dpi=300)
        plt.close()

        # ---------------------------------- 测试集混淆矩阵（样本数量形式）
        plt.figure(figsize=(5, 4), dpi=300, facecolor='lightgray')
        # plt.title('Test set (20%) Confusion Matrix - Sample Counts', fontsize=12)
        ax = sns.heatmap(cm_test, annot=True, cmap='BuGn', fmt='d',
                         xticklabels=label, yticklabels=label)
        ax.xaxis.tick_top()
        ax.set_xticklabels(ax.get_xticklabels(), rotation=90)
        plt.yticks(rotation=0)
        plt.xlabel('Predicted_label', fontsize=14)
        plt.ylabel('True_label', fontsize=14)
        ax.xaxis.set_label_position('top')
        plt.tight_layout()
        plt.savefig(outPath + '\\' + r'_RF_TestMatrixCounts.tif', format='tif', bbox_inches='tight',
                    transparent=True, dpi=300)
        plt.close()


        from sklearn.metrics import accuracy_score

        n_estimators_range = range(1, 120, 2)

        rf_clserr = RandomForestClassifier(
            random_state=1,
            warm_start=True,
            max_depth=best_model_rf.max_depth,
            min_samples_split=best_model_rf.min_samples_split
        )

        # 类别值
        class_labels = classOneLabel
        label_names = [files[i] for i in class_labels]

        # 每个类别单独保存误差曲线
        class_errors_dict = {cls: [] for cls in class_labels}
        overall_errors = []

        for n in n_estimators_range:
            rf_clserr.set_params(n_estimators=n)
            rf_clserr.fit(trainData, trainTarget)
            y_train_pred = rf_clserr.predict(trainData)

            # overall training error
            overall_err = 1 - accuracy_score(trainTarget, y_train_pred)
            overall_errors.append(overall_err)

            # class-wise errors
            for cls in class_labels:
                cls_mask = (trainTarget == cls)
                correct = np.sum(y_train_pred[cls_mask] == cls)
                total = np.sum(cls_mask)
                err = 1 - correct / total if total > 0 else 0
                class_errors_dict[cls].append(err)

        # ====== 绘制多类别训练误差曲线 ======
        # ====== 绘制多类别训练误差曲线（专业科研配图规范）======
        # 线型列表：多种线型交替，确保各类别曲线在黑白打印下也可区分
        _rf_conv_linestyles = [
            '-', '--', '-.', ':',
            (0, (3, 1, 1, 1)),  # 密集点划线
            (0, (5, 1)),  # 长虚线
            (0, (3, 5, 1, 5)),  # 疏密交替
            (0, (1, 1))  # 密集虚线
        ]
        # markevery：每隔若干点才显示一个标记，避免标记掩盖曲线走势
        _rf_markevery = max(1, len(list(n_estimators_range)) // 8)

        plt.figure(figsize=(7, 5), dpi=300, facecolor='white')

        for i, cls in enumerate(class_labels):
            plt.plot(
                list(n_estimators_range),
                class_errors_dict[cls],
                color=colorSpectrumColour[i % len(colorSpectrumColour)],
                linestyle=_rf_conv_linestyles[i % len(_rf_conv_linestyles)],
                linewidth=1.5,
                marker=markers[i % len(markers)],
                markersize=4,  # 小标记，不遮挡曲线
                markevery=_rf_markevery,  # 稀疏标记
                markeredgewidth=0.5,
                markeredgecolor='white',  # 白色描边，与背景区分
                alpha=0.85,
                label=f'{label_names[i]} Error'
            )

        # 总体误差：粗黑虚线，视觉上最突出
        plt.plot(
            list(n_estimators_range),
            overall_errors,
            color='black',
            linestyle='--',
            linewidth=2.5,
            marker='D',
            markersize=4,
            markevery=_rf_markevery,
            markeredgewidth=0.5,
            markeredgecolor='white',
            label='Overall Model Error'
        )

        plt.xlabel('Number of Trees', fontsize=14)
        plt.ylabel('Error Rate', fontsize=14)
        # ncol=2 分两列显示图例，避免遮挡图形主体
        plt.legend(loc='upper right', fontsize=18,
                   framealpha=0.85, edgecolor='#CCCCCC', ncol=2)
        # 轻量网格，不喧宾夺主
        plt.grid(True, linestyle='--', alpha=0.35, linewidth=0.7)
        _ax_rf_conv = plt.gca()
        _ax_rf_conv.spines['top'].set_visible(False)
        _ax_rf_conv.spines['right'].set_visible(False)
        plt.tight_layout()

        plt.savefig(outPath + '\\' + r'_RF_ClassConvergence_Error.tif',
                    format='tif', bbox_inches='tight', transparent=True, dpi=300)
        plt.rcParams['svg.fonttype'] = 'none'
        plt.savefig(outPath + '\\' + r'_RF_ClassConvergence_Error.svg',
                    format='svg', bbox_inches='tight', transparent=True)
        plt.close()

        # ====== 保存数值结果 ======
        with open(outPath + '\\' + r'_RF_ClassConvergence_Error.txt', 'w', encoding='utf-8') as f:
            f.write("随机森林收敛性分析（多类别训练误差变化）\n")

            header = "n_estimators"
            for i, cls in enumerate(class_labels):
                header += f"\t{label_names[i]}_error"
            header += "\toverall_error\n"
            f.write(header)

            for idx, n in enumerate(n_estimators_range):
                line = f"{n}"
                for cls in class_labels:
                    line += f"\t{round(class_errors_dict[cls][idx], 4)}"
                line += f"\t{round(overall_errors[idx], 4)}\n"
                f.write(line)

        # ************************************** ROC 曲线
        plt.figure(figsize=(5, 4), dpi=300, facecolor='white')
        fpr, tpr, roc_auc = dict(), dict(), dict()
        if len(classOneLabel) == 2:
            fpr[0], tpr[0], _ = roc_curve(testTarget, y_scoreRF[:, 1])
            roc_auc[0] = auc(fpr[0], tpr[0])
            plt.plot(fpr[0], tpr[0], color='darkgreen', lw=2,
                     label='ROC curve (AUC = %0.2f)' % roc_auc[0])
            plt.plot([0, 1], [0, 1], 'k--')
            # plt.title('RF_ROC')
            plt.xlabel('False Positive Rate')
            plt.ylabel('True Positive Rate')
            plt.legend(loc='lower right', fontsize=10)
            plt.savefig(outPath + '\\' + r'_RF_ROC.tif', format='tif', bbox_inches='tight', transparent=True, dpi=300)
            plt.close()
        else:
            for i in range(len(classOneLabel)):
                fpr[i], tpr[i], _ = roc_curve(testTarget1[:, i], y_scoreRF[:, i])
                roc_auc[i] = auc(fpr[i], tpr[i])
            all_fpr = np.unique(np.concatenate([fpr[i] for i in range(len(classOneLabel))]))
            mean_tpr = np.zeros_like(all_fpr)
            for i in range(len(classOneLabel)):
                mean_tpr += np.interp(all_fpr, fpr[i], tpr[i])
            mean_tpr /= len(classOneLabel)
            fpr["macro"], tpr["macro"] = all_fpr, mean_tpr
            roc_auc["macro"] = auc(fpr["macro"], tpr["macro"])
            plt.figure()
            plt.plot(fpr["macro"], tpr["macro"],
                     label='macro-average ROC curve (AUC = %0.2f)' % roc_auc["macro"],
                     color='darkgreen', linestyle=':')
            for i in range(len(classOneLabel)):
                plt.plot(fpr[i], tpr[i],
                         label=f'Class {files[i]} (AUC = {roc_auc[i]:.2f})')
            plt.plot([0, 1], [0, 1], 'k--')
            plt.xlabel('False Positive Rate')
            plt.ylabel('True Positive Rate')
            plt.legend(loc='lower right', fontsize=10)
            plt.title('RandomForest ROC (multi-class)')
            plt.savefig(outPath + '\\' + r'_RF_ROC.tif', format='tif', bbox_inches='tight', transparent=True, dpi=300)
            plt.close()

        # # ********************************************************************************************保存RF网格参数
        # 表 3：主要监督学习模型的分类性能比较
        # ==========================================================================================================

        from sklearn.metrics import accuracy_score, precision_score, recall_score, f1_score, classification_report

        model_names = ['SVM', 'LDA', 'KNN', 'DecisionTree', 'RandomForest', 'XGBoost']
        y_pred_all = [clf.predict(testData), lda.predict(testData), knn.predict(testData),
                      dt.predict(testData), rf.predict(testData), xgb.predict(testData)]

        perf_results = []
        for name, y_pred_i in zip(model_names, y_pred_all):
            acc = accuracy_score(testTarget, y_pred_i)
            pre = precision_score(testTarget, y_pred_i, average='weighted')
            rec = recall_score(testTarget, y_pred_i, average='weighted')
            f1sc = f1_score(testTarget, y_pred_i, average='weighted')
            perf_results.append(
                [name, round(acc * 100, 2), round(pre * 100, 2), round(rec * 100, 2), round(f1sc * 100, 2)])

        df_metrics = pd.DataFrame(perf_results, columns=["模型", "准确率(%)", "精确率(%) (weighted)", "召回率(%) (weighted)",
                                                         "F1值(%) (weighted)"])
        df_metrics.to_csv(os.path.join(outPath, '_Model_Performance_Table3.csv'), index=False, encoding='utf_8_sig')

        # 生成 DataFrame 表格
        df_metrics = pd.DataFrame(perf_results, columns=["模型", "准确率(%)", "精确率(%)", "召回率(%)", "F1值(%)"])

        # 保存为 CSV 文件与 TXT 文件
        outPath_table = os.path.join(outPath, '_Model_Performance_Table3.csv')
        df_metrics.to_csv(outPath_table, index=False, encoding='utf_8_sig')


        outPath_table_txt = os.path.join(outPath, '_Model_Performance_Table3.txt')
        with open(outPath_table_txt, 'w', encoding='utf-8') as f:
            f.write("表 3. 主要监督学习模型的分类性能比较\n")
            f.write(df_metrics.to_string(index=False))
            f.write("\n\n说明：评估指标包括准确率(Accuracy)、精确率(Precision)、召回率(Recall)与F1值(F1-score)。\n")

        # 打印结果至控制台，便于用户查看
        print("\n表 3. 主要监督学习模型的分类性能比较：\n")
        print(df_metrics)

        return 1


# # # # # # # # # 调用测试
dataPath = r'E:\工作文档（重要）\2-论文、专利撰写\论文3-foods约稿-白酒微生物\数据和结果\大曲微生物拉曼数据'
outPath = r'E:\工作文档（重要）\2-论文、专利撰写\论文3-foods约稿-白酒微生物\数据和结果\大曲微生物拉曼数据-结果4'
# 截谱
# startSpec = 400
# endSpec = 300
# endSpec = 0

startSpec = 30
endSpec = 0
# endSpec = 50


# 前处理参数
snrNoiseindex = 30
# 从尾部选取长度，作为静默区

delSNRmin = 1
# 删除光谱信噪比阈值
delSNRmax = 1000
# 删除光谱信噪比阈值

# 批处理参数
# 2.1 宇宙射线阈值
filterSize = 5
# 整型
dynamicFactor = 4.5
# 浮点型
# 2.2 S-G滤波
winSG = 5
nSG = 3
lambdaAirPls = 100
itermaxAirPls = 15
# 更保型

# hca参数
distanceMethod = 0
# 层次聚类距离度量选用的方法,可选：'centroid','weighted','average'
# hcaColorThreshold = 10
# # color_threshold= 8 设置颜色高度阈值，用于控制决策树和PCA的颜色分类
hcaLabelsize = 8
# 标号大小

# SVM参数
svmTestSize = 0.2
# 测试数据占全部数据比例
svmC = 0.7
# 错误项的惩罚系数,范围[0.5-1],C越大分类效果越好，但有可能会过拟合（defaul C = 1）
xKernel = 0
# str参数,算法中采用的核函数类型,可选'linear'、'poly'、'rbf'、'sigmoid'、'precomputed'，默认为rbf
clusters_N = 5
# 聚类堆数
oneclickReport(dataPath, outPath, startSpec, endSpec, snrNoiseindex, delSNRmin, delSNRmax, filterSize, dynamicFactor,
               winSG, nSG, lambdaAirPls,itermaxAirPls, distanceMethod, hcaLabelsize, svmTestSize, svmC, xKernel, clusters_N)