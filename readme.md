# 绩效考核

## 一.依赖

### python 环境

一定要安装 3.x 版本, 2.x 环境上没有测试过

### 读取 excle 库

```shell
pip install pandas

pip install openpyxl
```

## 二.如何使用

1. 把所有的 xlsx 文件放到和 main.py 同一级目录下
2. 运行

```shell
python main.py
```

3.运行结果会在目录下生成一个 report.xlsx 文件

## 需求

### bug 提交

- 先按照分界线， 比如 15
- 平均月 BUG 小于 15 只有 40 分
- 剩下的排序 大于 15 的需要排名
  - 第一名 100 分
  - 第二名分数等于: (自己的 bug 数 / 第一名的 bug 数) \* 100

### 文档评审

ireview 和比去年比较, 分为半年和全年

- 本年业绩相比去年，增长小于 30%， 则 40 分
- 否则
  - 增长最多的，第一名 100 分
  - 第二名: 自己的 (review 数 / 第一名 review 数) \* 100
