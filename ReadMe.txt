1、本文件是将指定目录下面的所有java文件合并到一个word文档。然后把根据指定目录下面的目录结构设计word目录等级。
2、打包成window可执行文件方法。
2.1 使用管理员权限执行，安装pip install pyinstaller
2.2 找到可执行文件pyinstaller.exe 
2.3 打包：D:\AllServer\Anaconda\Scripts\pyinstaller.exe  --onefile main.py --name JavaFileMerge.exe
方法二：执行build.py配置文件 python  build.py
	

打包遇见：
The 'pathlib' package is an obsolete backport of a standard library package and is incompatible with PyInstaller. Please remove this package (located in D:\AllServer\Anaconda\lib\site-packages) using
    conda remove
then try again.
解决方法：conda remove pathlib
执行失败，配置国内镜像源
conda config --set show_channel_urls yes
conda config --add channels https://mirrors.tuna.tsinghua.edu.cn/anaconda/pkgs/free/
conda config --add channels https://mirrors.tuna.tsinghua.edu.cn/anaconda/pkgs/main/

手动配置： 用户目录下\.condarc 文件，添加下面的内容
channels:
  - https://mirrors.tuna.tsinghua.edu.cn/anaconda/pkgs/main/
  - https://mirrors.tuna.tsinghua.edu.cn/anaconda/pkgs/free/
