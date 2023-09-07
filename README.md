[![Anurag's github stats](https://github-readme-stats.vercel.app/api?username=evilbutcher)](https://github.com/anuraghazra/github-readme-stats)

# [文献下载小程序](https://github.com/evilbutcher/Python/tree/master/ArticlesHelper)

[English Version](https://github.com/evilbutcher/Python/blob/master/README_EN.md)

一开始写了 JavaScript 版的[文献下载助手](https://github.com/evilbutcher/Code/tree/master/%E6%96%87%E7%8C%AE%E4%B8%8B%E8%BD%BD/%E6%96%87%E7%8C%AE%E4%B8%8B%E8%BD%BD%E5%8A%A9%E6%89%8B)，但这个只能在 JSBox 上运行，有一定的限制和门槛。

时至今日，我终于捡起来 Python，开始着手移植，一边移植一边学 python...

#### 如果实际使用下载速度太慢推荐"加速器"：[Matrix](https://amatrixap.com/auth/register?code=UFMM) 、[GLaDOS](https://glados.space/landing/3JRG4-KSGZJ-8QPXF-8PPOO)（3JRG4-KSGZJ-8QPXF-8PPOO）

### 关于如何使用

请前往[Releases](https://github.com/evilbutcher/Python/releases)

#### For Windows:

下载最新的文献下载助手小程序.exe。

#### For macOS:

因为没有 Mac，所以提供了源码，请自行测试，用到的模块有：

```python
import requests
import os
import re
from rich.console import Console
from rich import print
from rich.table import Table
from pathlib import Path
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor
from functools import partial
from urllib.request import urlopen
from rich.progress import (
    BarColumn,
    DownloadColumn,
    TextColumn,
    TransferSpeedColumn,
    TimeRemainingColumn,
    Progress,
    TaskID,
)
```

首次运行，会在同级目录生成两个文件夹，一个是 articles，用于存储下载的文献，另一个是 records，用于存储下载的 Web of Science 文献记录。

#### 如何下载 Web of Science 文献记录

选中文献后，导出格式中选择 html，然后将下载的 html 文件转存到同级目录的 records 文件夹中。

#### 如何解析记录

如果程序检测到在 records 中，存在.html 格式的文件，就会自动将名称列出来，提示是否进行解析，输入 y 则会执行解析，n 则会返回手动输入 doi 号下载。

如果存在 records 中存在多个 html 文件，可直接输入全称如 savedrecs.html 进行解析，不解析请输入 n 。


#### 如何手动下载

直接输入 doi 号即可下载，多个 doi 请用英文逗号“,”进行分割，例如 10.1016/j.snb.2013.07.010,10.1016/j.snb.2010.12.010,10.1039/c5cs00424a。


#### 自动检测更新

如果有更新，软件会自动弹出更新提示，可前往[Releases](https://github.com/evilbutcher/Python/releases)地址进行更新。


### 现已支持

1. 根据 doi 进行文献下载和保存
2. 下载异常判断
3. 批量下载
4. 自动交替请求下载
5. 下载失败自动更换地址
6. 进度条
7. 自动检测更新
8. 解析 Web of Science 文献记录
9. 下载情况检查
10. 预先检查 articles 文件夹

# [基线拉平小程序](https://github.com/evilbutcher/Python/blob/master/evanescent)

处理实验数据自用

# [Q-PCR数据处理](https://github.com/evilbutcher/Python/blob/master/Q-PCR)

处理实验数据自用

### 特别感谢：

[@Rich](https://github.com/willmcgugan/rich)

### 访问量

![](http://profile-counter.glitch.me/evilbutcher/count.svg)
