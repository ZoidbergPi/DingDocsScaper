# DingDocsScaper
钉钉文档爬取+下载工具

## 免责申明
本工具仅供技术研究以及交流，禁止用于真实的钉钉文档导出场景，所有使用脚本的个人需要自己承担任何可能存在的的法律风险，本仓库概不负责。

## 测试方式
1、安装Python 3.6版本以上

2、安装依赖： pip -r requirements.txt

3、修改脚本中的参数，公司ID corpId 以及需要爬取的知识库（组织）ID，也就是这个URL：https://alidocs.dingtalk.com/i/spaces/知识库ID/overview?corpId=公司ID 打开后能够看到你要爬取的内容才可以

4、执行脚本 python main.py

5、等待依赖部署好浏览器，并打开浏览器后，在控制台按任意键开始爬取



## 限制
1、目前仅支持下载文档以及表格，其它的可以自己根据process_file()方法中的逻辑进行补充


## 使用到的第三方库
loguru、DrissionPage[https://github.com/g1879/DrissionPage]

## 鸣谢
感谢leisurelyclouds/DingTalkSpaceExporter 提供思路
