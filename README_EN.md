[![Anurag's github stats](https://github-readme-stats.vercel.app/api?username=evilbutcher)](https://github.com/anuraghazra/github-readme-stats)

# [ArticlesHelper](https://github.com/evilbutcher/Python/tree/master/ArticlesHelper)

#### Recommended "Accelerator" if the actual download speed is too slow: [Matrix](https://amatrixap.com/auth/register?code=UFMM) 、[GLaDOS](https://glados.space/landing/3JRG4-KSGZJ-8QPXF-8PPOO)（3JRG4-KSGZJ-8QPXF-8PPOO）

### How to Use

1. Please go to the [Releases](https://github.com/evilbutcher/Python/releases)

#### For Windows:

Download the latest ArticlesHelper.exe。

#### For macOS:

Since I don't have a Mac, I've provided the source code, please test it yourself.

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

The first run generates two folders in the same level directory, one is articles, which stores downloaded articles, and another is records, which stores downloaded Web of Science articles records.

#### How to download the Web of Science records

After selecting articles, select html in the export format, and then save the downloaded html file to the records folder in the same level directory.


#### How to parse records

If the program detects the existence of a .html file in records, it automatically lists the names and prompts whether to parse it or not.

If there are multiple html files in records, you can enter the full name such as savedrecs.html for parsing. If you do not want to parse it, input n.


#### How to download manually

Enter the doi number directly to download, multiple doi can be divided by comma ",", e.g. 10.1016/j.snb.2013.07.010,10.1016/j.snb.2010.12.010,10.1039/ c5cs00424a.


#### Automatic detection of updates

If there is an update, the software will automatically pop up an update prompt, you can go to the [Releases](https://github.com/evilbutcher/Python/releases) to update.



### Now supported

1. Download and save documents according to doi
2. Determine download anomalies
3. Batch downloads
4. Automatic change request downloads
5. Automatic address change for failed downloads
6. Progress bars
7. Automatic detection of updates
8. Parsing the Web of Science documentation
9. Checking of downloads
10. Pre-check the articles folder

# [Baseline Alignment](https://github.com/evilbutcher/Python/blob/master/evanescent)

Processing experimental data for personal use

# [Q-PCR](https://github.com/evilbutcher/Python/blob/master/Q-PCR)

Processing experimental data for personal use

### Special thanks to

[@Rich](https://github.com/willmcgugan/rich)

### Visitors

![](http://profile-counter.glitch.me/evilbutcher/count.svg)
