- 参考[文章](https://www.python.jp/install/anaconda/index.html)安装 [Anaconda3-YYYY.MM-Windows-x86_64.exe](https://www.anaconda.com/products/individual)。

- 删除 main.py 中的操作系统相关代码（含 import 部分），

  只保留 Windows 相关代码，以尽量减小导出的可执行文件大小。

- 执行导出代码。

  参考：[别再问我Python打包成exe了！（终极版）](https://mp.weixin.qq.com/s?src=11&timestamp=1610821341&ver=2832&signature=29SfQcuHuQcpekH8HBqn5Ltma4NUqnzifMxFpqIoO70HAaoO8op1A55D7geHgIsoVRkI5zoqD0i30xJXsMYllN5vJVb8bQLBttx0*Izmi12PpUp4SwlnwAHGAcXTjHR5&new=1)

  ```shell
  #创建虚拟环境
  conda create -n test python=3.8
  #激活虚拟环境
  conda activate test
  #查看虚拟环境
  conda info -envs
  #查看当前虚拟环境里已安装的库
  conda list
  #安装依赖
  pip install pyinstaller
  pip install openpyxl
  #开始打包
  pyinstaller -F -w main.py
  #退出虚拟环境
  conda deactivate
  #删除虚拟环境
  conda remove -n test --all
  ```

- 将 `input/` `output/` `sample/` `sqlite/` `template/` 和 `main.exe` 打包成压缩文件。

  > **ToukiGassen-1.0.0-Windows-x86_64.zip**

