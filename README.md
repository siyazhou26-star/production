# 组会汇报 PPT 工作区

把研究材料、实验图表、汇报时长和学校模板交给 Codex，即可协助整理内容并生成可编辑的 PowerPoint 文件。不需要编程经验或 API 密钥。

## 环境与使用

Python 3.12。安装：

```sh
cd /workspace/production
python -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python make_ppt.py examples/report.json outputs/组会汇报模板.pptx
.venv/bin/python -m unittest discover -s tests -v
```

`examples/report.json` 是中文七页模板，所有内容均为占位提示，不含真实研究结论。生成的 PPT 为 16:9，文字可编辑，可用 PowerPoint 或 WPS 打开并插入图表。每页建议 3–5 条简短要点，长内容请拆页；此基础模板不自动排版图片。

生成文件放在 `outputs/`，默认不提交到 Git，以免上传未公开的研究数据。中文字体使用微软雅黑；打开文件的电脑若缺少该字体会使用替代字体。导出程序不会调用外部 AI 服务或上传研究资料。

后续只需告诉 Codex：“根据这些材料制作一份 10 分钟的组会汇报 PPT”，并提供材料即可。实际汇报应核对数据、结论和引用。
