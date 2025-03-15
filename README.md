# ProgressBar for PowerPoint
PowerPointのVBAマクロでスライドの進行状況バーを生成します。
![image](https://user-images.githubusercontent.com/78206853/209463476-e071bc7f-92eb-4b94-ad03-6bfc0f1a0923.png)

### Usage
1. 適当な名前(ProgressBar)を付けてマクロを作成する
2. [ProgressBar.bas](ProgressBar.bas)の内容をコピー&ペースト
3. 実行
4. 全画面の切替効果を「変形」にする  
  スライドが増減した際は、再度実行

### Properties
変数名|値
---|---
objectName|図形の名前
barHeight|バーの縦幅
offsetVertical|バーの下からの位置
barWidth|バーの横幅
offsetHorizontal|バーの左からの位置
barColor|バーの色(RGB)
