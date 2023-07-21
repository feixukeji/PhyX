import os
import chardet
import traceback
import pandas as pd
import math
import numpy as np
import scipy.optimize
import scipy.stats
import matplotlib
import matplotlib.pyplot as plt
from docx import Document
from docx.oxml.ns import qn
from api.calc import *
from api.insert import *

fig, ax = plt.subplots() # 新建绘图对象
