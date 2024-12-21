# 操作指南


## 下载人员名单

从 BVF 下载人员名单（已加入和未加入分别命名为 `in.html` `out.html`）到 `csv/raw/` ，运行 `convert.py`，在 `csv/name list/` 中找到 `in.csv` `out.csv`，放入对应的文件夹

## 转换共享文档

保存共享文档到`xlsx/`，命名为年份+月份 (eg: `2024.9.xlsx`)，运行 `convert.py`，找到 `/converted/csv/` 下的csv文件（命名与sheet名相同），检查无误后放入 `csv/csv/`，将不同项目分类到同一个csv（可新建，命名随意，后面在代码中微调即可）

## 处理

将 `csv/csv/` 中已经处理好的csv文件和 `csv/name list/`中需要的名单放到 `csv/`上级目录下，然后打开 `BVFS.py`，微调csv命名，Run！

## 收集输出

打开 `output/`，三个文件内容如下：

|      文件名      |            内容             |
| :-----------: | :-----------------------: |
|  `find.txt`   |      找到的人，直接去BVF录入即可      |
| `result.txt`  | 没找到的人，写入 `data/data.txt`  |
| `result.xlsx` | 没找到的人，写入 `data/data.xlsx` |

最后，每隔一段时间去微信群发 `data/data.txt` 和对应项目的二维码（注明项目名称），去企业微信补充 `data/data.xlsx`

## 结构

```
BFVQS/
│
├── csv/
│   ├── csv/
│   └── name list/
│   └── raw/
│        ├── converted.py
│        └── in.html
│        └── out.html
│
├── data/
│   ├── data.txt
│   └── data.xlsx
│
└── output/
│   ├── find.txt
│   └── result.txt
│   └── result.xlsx
│
├── xlsx/
│   ├── converted/
│   │   └── csv/
│   └── convert.py
│
└── BFVQS.py
│
└── README.md
│
└── LICENSE
```

