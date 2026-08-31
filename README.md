# FastMoss 达人合并 · Streamlit V1

## 已配置渠道

- 商品关联达人：销量、销售额
- 商品关联视频：该商品销量、该商品销售额
- 带货达人榜：暂无核心指标
- 黑马达人榜：暂无核心指标
- 竞品店铺带货达人：[天]带货销量、[天]带货销售额
- 热门视频：暂无核心指标

## 本地运行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud

1. 将本目录中的文件推送到 GitHub 仓库。
2. 在 Streamlit Community Cloud 新建 App。
3. Repository 选择该仓库，Main file path 填 `app.py`。
4. Deploy。

## V1 处理逻辑

- 一次上传多个 xlsx / xls / csv。
- 主要依据表头结构识别渠道，不依赖文件名。
- 核心指标筛选默认关闭；启用后使用“≥ 最低值”。
- 若某个源表没有某指标字段，该指标不会作用于这个源表。
- 渠道内先筛选，再按达人ID去重。
- 有核心指标的渠道，重复达人优先保留核心指标更高的一条记录。
- 全渠道再次按达人ID去重，并合并“来源渠道 / 来源文件”。
- 缺少达人ID的记录无法去重，V1 会保留并在界面提示。
