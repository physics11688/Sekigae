# Sekigae
どこぞの高専の席替え用プログラム2022です。

<br>

# Requirement

以下で実行確認済み.
* Python 3.9 以降
* openpyxl 3.0.9 以降
* pandas 1.4.1 以降

```bash
$ pip install openpyxl pandas  # pandas と openpyxlのインストール
```

<br>

## Install
1. `$ git clone git@github.com:physics11688/Sekigae.git`

<br>

2. `$ cd Sekigae`

<br>

3. `$ python seating.py     # Mac or Linux --> python3 setup.py`

<br>


## Usage

ターミナルで `seating.py` を実行すると

1. `2022座席.xlsx` から名標データを読み込む
2. 名標データをシャッフルして座席表ファイル `座席_日付.xlsx` を吐き出す

みたいに動きます。

いくらでも機能はつけられるけど、1年が読めるようにできるだけ簡単にしてあります。


```bash
$ python seating.py      # シャッフルされた座席ファイルが出来る

$ python seating.py aaa  # 出席番号順の席ファイルが出来る.4月の初め用.
```

<br>
