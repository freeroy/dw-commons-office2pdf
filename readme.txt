1、转换器有两种模式实现
1）jacob模式，只能运行在window环境，且需要对应安装office软件
2）jod模式（openoffice模式），能在linux、window环境运行，需要安装openoffice软件
3）在window场景下，使用jacob模式性能会更佳

2、jacob模式注意事项
1）需要在window操作系统中，安装office软件，实测office2007~office2013版本是通过的
2）需要下载对应的jacob dll文件，放置在java jdk安装目录下的jre/bin目录里面，dll有操作系统位数区分
3）dll下载地址：http://danadler.com/jacob/

3、jod模式注意事项
1）需要在操作系统中，安装openoffice软件
2）程序在运行时，会启动openoffice的转换器进程
3）JobConvertor虽然是线程安全，但每个JobConvertor都会启动一个openoffice的转换进程，因此建议在实际应用时，使用单例模式构建jobConvertor（即有且只有一个jobConvertor实例），减少由于进程过多而导致资源耗用，及提高转换运行速度
4)在用法上，jobConvertor会有点不同，在执行conver2Pdf方法前，必须显示调用一次startService方法，以便后台启动openoffice转换进程（启动后以后就不需要再启东）；同时不再使用jobConvertor时，也要执行stopService方法，关闭后台转换进程（在单例模式下，建议在容器关闭时一起关闭）

4、其它
暂时只支持word、powerpoint、excel转换pdf。