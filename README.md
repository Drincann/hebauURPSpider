# URP 爬虫

用于河北农业大学教务系统的爬虫。

学校的教务系统似乎没打某些升级补丁，有一个查询信息接口的权限没做验证。

为了随时拿到平均学分绩的排名，用 python 做了简单的封装。

依赖 request_html、pyquery、opencv-python。

> 偶然发现的 pyquery 包，可以以 jquery 的接口风格来处理 html，非常赞！

## 使用

```sh
pip install opencv-python pyquery request_html

python ./main.py
```
