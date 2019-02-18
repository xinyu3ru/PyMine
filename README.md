PyMine全自动扫雷脚本
-----------------


#### 依赖
<br>

- 依赖 pywin32 、 PIL

  安装

  `pip install -r requirements`

  `pip install pywin32 Pillow`

  如果pywin32安装失败，请直接安装.whl

  下载 https://sourceforge.net/projects/pywin32/files/pywin32/
- 需要自己下载扫雷，http://www.saolei.net/BBS/ ，右上角，下载。
- Win 10、Python 3.7

#### 使用

  1.打开扫雷

  2.运行本脚本，全自动扫雷

  3.Win 键强制退出脚本

#### 思路
  随机点击几处，截图分析界面情况，扫描各处到二维数组

  如果周围空白 + 红旗 == 提示数字，则空白均有雷，右键点击空白插红旗

  如果周围插旗 == 数字，则附近空白均没有雷，左键点击空白

  循环上面过程，如果没有符合条件的，则随机点击一个白块

#### >参考资料
  cox@swust 提供扫雷思路和C++源码和完美的开发文档。C++代码和软件







#### 感谢

- 感谢 cox@swust 提供扫雷思路和C++源码和完美的开发文档。
- 感觉各位网友在各个网页提供的资料。
