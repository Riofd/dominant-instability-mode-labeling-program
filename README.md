#  **一个图片样本标注辅助程序**
## **1.使用说明**
* 可以先用pyinstaller生成一个可执行exe文件，文件太大了，传不到github；推荐直接用pycharm运行代码
* 运行程序后弹出的第一个窗口选择样本文件夹
* 第二个窗口选择标签存放文件夹
## **2.说明**
* 样本文件夹需要包含的是图片，支持jpg,png等
* 标签文件只支持excel文件，请预先构建一个，并命名为label.xlsx
* 本程序可以自动识别标签文件中已标注的样本，并自动开始标注，如需要重新标注，请清除标签文件。
* 在菜单栏中“文件”下可以重新选择样本文件夹和标签文件夹
* 本程序默认exe程序执行的是主导失稳模式样本的标注 
### **注：需要提前新建一个空的excel文件用于存放标签，并命名为‘__label.xlsx__’**
## **3.Citation**
[A graph attention networks-based model to distinguish the transient rotor angle instability and short-term voltage instability in power systems](https://www.sciencedirect.com/science/article/pii/S0142061521010036?via%3Dihub)
```
@article{Zhang2021DIM,
  title={A Graph Attention Networks-Based Model to Distinguish the Transient Rotor Angle Instability and Short-term Voltage Instability in Power Systems},
  author={R. Zhang, W. Yao, Z. Shi, et al.},
  journal={International Journal of Electrical Power and Energy Systems},
  year={2022}
}
```
## 2020/9/16更新
* __修复了读取文档顺序的bug__
## 2021/12/30更新
* __修复了部分bug__
## __免责声明__
* **本程序完全开源，作者不对程序错误等造成的后果负责**
### <p align="right"> **Designed By ZRF from HUST SEEE**</p>