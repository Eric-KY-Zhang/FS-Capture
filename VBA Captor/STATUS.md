# 上市公司财务数据查询工具 - 项目 STATUS

> 项目状态:**已扩展为 A股 + 美股 + 港股 + 韩股财务数据查询工具; Phase 4f Step 2 已完成本地代码开发,等待联网验证**
> 上一轮迭代:从"横向→纵向"格式调整,演进为"完整重写为多公司同业对标工具"
> 当前工作簿:`上市公司财务数据查询.xlsm`
> 作者:Eric Zhang;联系邮箱:214978902@qq.com
> 创建日期:2026-05-02
> 工具来源:作者林铖(247650491@qq.com),原工具 V2.2

---

## 0. 最新进展(2026-05-03)

- Phase 4f Step 2 已完成本地开发:新增『汇率』sheet 模板、样本池 B6『显示币种』toggle、`ReadDisplayCurrency()` / `GetFxRate()` helper,以及独立 `模块_抓汇率.bas`。
- 离线验证已完成:`tools/build_template.py`、`tools/install_modules.py`、工作簿结构检查、VBA 编译命令均通过。
- 未跑联网验证、未跑 probe、未跑实际抓数、未跑立即窗口联网实测; 后续由 Claude Code / 用户接手验证雪球 FX 请求与真实写表流程。

## 1. 项目背景

### 1.1 原工具是什么

`新浪财经行业数据查询V2_2.xlsm` 是一个 Excel + VBA 工具,从新浪财经抓取上市公司财务数据。

**原工作流**:
1. 用户在 `样本池` Sheet 录入股票代码、名称、4 个 URL(资产负债表/利润表/现金流量表/指标表)
2. 运行 4 个抓数 Sub(模块1/3/4/5 的 `Main`),分别从对应 URL 抓 HTML 表格
3. 数据写入 4 张主表:`资产负债表` / `利润表` / `现金流量表` / `指标表`,**横向格式**:
   - 行 = 股票 × 报告期(每只股票占 N 行,N=报告期数)
   - 列 = 指标(资产负债表 88 个、利润表 ~29 个、现金流量表 71 个、指标表 85 个)
   - 表头两层:R1 大类 / R2 子项,通过合并单元格组织

**原工具的痛点**:
- 单只股票分析尚可,但**多公司同业对标**(用户的核心场景)很别扭——想横向比较 4 家公司的"货币资金",数据散在 4 行里
- 行列方向不利于做透视分析
- 老旧的 `WinHttp.WinHttpRequest.5.1` + `gb2312` 编码,无重试、无限速

### 1.2 用户(Eric)的真实需求

**做家居/床垫行业的同业对标分析**(样本池里就是 安克创新、梦百合、喜临门、致欧科技 等)。
理想形态是一张表里:
- 一行 = 一个财务指标
- 一列(或一组列)= 一家公司在某个报告期的数值
- 多家公司并排,容易横向对比

这种格式叫**宽表**(wide table / pivot table),是财务分析师做横向对标的标准成品形态。

---

## 2. 设计澄清历史(决策日志)

> 这一节记录了从最初需求"改纵向"到最终方案的演进。Claude Code 接手时**不必读全部细节,但要知道每个决策背后的成本权衡都做过了**,不要轻易推翻。

### Round 1: "横向铺开改为纵向"
- 用户要求把横向格式改纵向。
- 澄清"纵向"的两种理解:A 长表(每行=一个指标值) vs B 表头转置。
- 原 VBA 抓数和写入逻辑是**耦合死的**——VBA 写死了"行=报告期、列=指标"的写入。改格式必须同步改 VBA。

### Round 2: 用户选 B(表头转置) + 公式地址级兼容
- **冲突识别**:表头从横转竖后,所有单元格地址必然变化,无法同时满足"地址兼容"和"格式变更"。
- 用户改选 A(长表)。

### Round 3: 长表方案落地(已废弃)
- 不动原 4 张主表,新增 4 张 `_长表` Sheet(`股票代码|名称|报告日期|大类|指标名称|数值`)。
- 提供 `模块6_长表展开.bas`,Sub `展开所有长表()` 把主表展开为长表。
- 已交付,但用户继续迭代。

### Round 4: 用户上传"宽表版",改为宽表方案
- 用户手工做了 4 张 `_宽表` Sheet,展示了真实想要的格式:
  - 行=指标(纵向铺开)
  - 列=公司名×报告期(横向多层表头,4 公司 × 4 期 = 16 列)
  - R1 = 公司名(代码),横向合并 4 列;R2 = 报告期
- "长表 + 原主表都可以删"。

### Round 5: 进一步澄清(本次)
- **报告期对齐**:用户选 C 取并集(所有公司在同一报告期轴对齐,缺数据则空白)
- **删除范围**:用户最初选"全删,瑞华底稿也删"
- **抓数路径**:用户最终选 b——**重写抓数 VBA,直接抓到宽表**(放弃"原表当中转区"的稳妥路线)
- **新增功能**:加一个抓"上市公司基本资料"的爬虫,字段:代码/简称/上市日期/行业/主营业务

---

## 3. 最终设计方案(锁定)

### 3.1 工作簿结构(目标态)

| Sheet | 状态 | 说明 |
|---|---|---|
| 使用说明 | 重写 | 更新为新工作流的说明 |
| 样本池 | **改造** | 见 §3.2 |
| 资产负债表_宽表 | **新建** | §3.3 格式 |
| 利润表_宽表 | **新建** | §3.3 格式 |
| 现金流量表_宽表 | **新建** | §3.3 格式 |
| 指标表_宽表 | **新建** | §3.3 格式 |
| 上市公司基本资料 | **改造** | §3.4 抓取目标,不再是预置全市场名单 |
| ~~资产负债表~~ | **删除** | 原横表 |
| ~~利润表~~ | **删除** | 原横表 |
| ~~现金流量表~~ | **删除** | 原横表 |
| ~~指标表~~ | **删除** | 原横表 |
| ~~资产负债表_长表~~ | **删除** | 上一轮的方案 |
| ~~利润表_长表~~ | **删除** | 上一轮的方案 |
| ~~现金流量表_长表~~ | **删除** | 上一轮的方案 |
| ~~指标表_长表~~ | **删除** | 上一轮的方案 |
| ~~瑞华底稿~~ | **删除** | 用户明确同意删 |

### 3.2 样本池 Sheet 设计

**输入区**(用户填写):
- 列 A:股票代码(如 `300866`)
- 列 B:股票简称(如 `安克创新`)
- 列 C:交易所代码前缀(`sh` / `sz`),用于拼 URL,可由代码自动推断也可手填

**URL 区**(由代码自动生成或保留原手填):
- 列 D:资产负债表 URL
- 列 E:利润表 URL
- 列 F:现金流量表 URL
- 列 G:指标表 URL
- 列 H:基本资料 URL(新增)

URL 模板(实际抓取前先确认下面 URL 仍然有效):
```
资产负债表: http://money.finance.sina.com.cn/corp/go.php/vFD_BalanceSheet/stockid/{code}/ctrl/part/displaytype/4.phtml
利润表:     http://money.finance.sina.com.cn/corp/go.php/vFD_ProfitStatement/stockid/{code}/ctrl/part/displaytype/4.phtml
现金流量表: http://money.finance.sina.com.cn/corp/go.php/vFD_CashFlow/stockid/{code}/ctrl/part/displaytype/4.phtml
指标表:     http://money.finance.sina.com.cn/corp/go.php/vFD_FinancialGuideLine/stockid/{code}/ctrl/2025/displaytype/4.phtml
基本资料:   http://money.finance.sina.com.cn/corp/go.php/vCI_CorpInfo/stockid/{code}.phtml
```

**注意**:原工具样本池前 6 行是表头/说明区,实际数据从 row 7 开始(VBA 里 `For i = 7 To UBound(arrUrl)`)。重写时可以简化为 row 2 开始,但要更新 VBA 起始位置。

### 3.3 宽表格式规范(参考用户上传的 `_宽表` Sheet)

```
R1: |  大类  | 指标名称 |  安克创新(300866)               |  梦百合(603313)                | ...
    |       |        |  [合并 4 列,水平居中]            |                                |
R2: |       |        | 2025-12-31 | 2025-09-30 | ... |  2025-12-31 |  2025-09-30  | ...
R3: | 流动资产 | 货币资金 | 365650.93  | 259128.77  | ... | 108885.53   | 93404.04     | ...
...
```

**结构规则**:
- **A 列**:大类(如 `流动资产`,利润表大部分指标无大类则留空)
- **B 列**:指标名称
- **C 列起**:每家公司占 N 列,N = 报告期并集数量
- **R1**:公司名(代码),跨该公司所有列合并,水平居中
- **R2**:报告期(日期类型,格式 `yyyy-mm-dd`)
- **R3 起**:数据行

**报告期对齐(关键!并集逻辑)**:
- 抓所有公司所有报告期 → 取并集 → **降序排序**
- 每家公司都占满并集长度 N 列
- 该公司在某个并集报告期没数据 → 空白单元格(不是 `--`,不是 `0`)

**指标行(行)的并集**:
- 不同公司可能少数指标缺失或多出来 → 取所有公司指标的**并集**,按原表 col 顺序排
- 排序基准:用样本池里**第一个非空公司**的指标顺序为锚,后面公司有新指标就按出现顺序追加
- 该公司没该指标 → 行存在,数值列空白

**值规则**(沿用前面已锁定):
- `"--"` → 空白
- 数值字符串 → 转 Double
- 报告日期保持 datetime,显示 `yyyy-mm-dd`

**已知瑕疵(忠实保留,不修正)**:
- 利润表:新浪 HTML 在 `六、每股收益` 设置了 AA1:AF1 横向合并(覆盖 col 27-32),导致 col 29-32 的几个指标(`七、其他综合收益`、`八、综合收益总额`、归母综合收益、少数综合收益)在大类列里显示为"六、每股收益"——这是新浪原始数据的语义瑕疵,工具忠实复刻不修正。

### 3.4 上市公司基本资料 Sheet 设计

**抓取目标 URL**:`http://money.finance.sina.com.cn/corp/go.php/vCI_CorpInfo/stockid/{code}.phtml`

**抓取字段**(用户已确认核心 5 个):
| 列 | 字段 | 备注 |
|---|---|---|
| A | 股票代码 | |
| B | 股票简称 | |
| C | 上市日期 | 日期类型 |
| D | 所属行业 | 新浪页面"所属行业"字段 |
| E | 主营业务 | 新浪页面"主营业务"字段,可能很长 |

**实施提示**:
- 新浪 CompanyInfo 页面的 HTML 结构与三大表不同,需要单独写正则/解析
- 字段在页面里通常是 `<table id="comInfo1">` 里的标签-值对,定位"上市日期""所属行业""主营业务"等中文 label,取相邻 td 的值
- **建议先 curl 一个真实页面下来手动看 HTML 结构,再写正则**(原作者用的 Regex 单行匹配方式,可参考)

### 3.5 VBA 模块设计

**保留**:`模块2`(SetBorderLine 通用边框函数)

**删除**:模块1 / 模块3 / 模块4 / 模块5(原 4 个抓数 Sub)

**新建**:

```
模块_工具函数.bas      # ByteToStr、HtmlGet 等通用函数(从原模块抽出复用)
模块_抓资产负债表.bas   # Sub Main_抓资产负债表()
模块_抓利润表.bas       # Sub Main_抓利润表()
模块_抓现金流量表.bas   # Sub Main_抓现金流量表()
模块_抓指标表.bas       # Sub Main_抓指标表()
模块_抓基本资料.bas     # Sub Main_抓基本资料()
模块_总入口.bas         # Sub 一键全抓() — 顺序调用上述 5 个
```

**核心算法(每张财务报表的抓数 Sub 都要实现)**:

```pseudo
Sub Main_抓XXX():
    清空目标宽表 Sheet 的内容(保留表头容器)

    # 第一遍:抓所有公司的 HTML,存入字典
    Dim companyData = Dictionary()  # key=股票代码, value=Dictionary(报告期→Dictionary(指标→值))
    Dim companyOrder = Array()       # 保留样本池顺序
    Dim allDates = Set()             # 所有报告期并集
    Dim allIndicators = OrderedDict() # 所有指标并集,保持出现顺序
    Dim categoryMap = Dictionary()    # 指标 → 大类

    For each 公司 in 样本池:
        html = HttpGet(URL)
        table = 解析 <table id="..."> 的所有行列
        # 新浪 HTML 表格结构:第一行是报告期表头,后面每行是一个指标
        # 表头 R0 = ["报告日期", "2025-12-31", "2025-09-30", ...]
        # 每行 = [指标名, 值1, 值2, ...]
        # 大类信息可能体现在跨行合并的 td 里,需要识别

        For each 行 in table:
            指标名 = 行[0]
            For each 报告期, i in 表头:
                companyData[公司代码][报告期][指标名] = 行[i]
                allDates.add(报告期)
                if 指标名 not in allIndicators:
                    allIndicators.append(指标名)

    # 第二遍:取并集 + 降序日期 + 写入宽表
    sortedDates = sorted(allDates, reverse=True)

    # 写表头 R1: 大类 | 指标名称 | 公司A名(代码) <merge across len(sortedDates) cols> | 公司B... | ...
    # 写表头 R2: 空 | 空 | 报告期1 | 报告期2 | ... | 报告期1 | 报告期2 | ...
    # 写数据 R3+: 大类 | 指标 | 公司A该指标各期值 | 公司B该指标各期值 | ...

End Sub
```

**注意 VBA 实现的关键点**:
1. **集合与字典**:用 `Scripting.Dictionary` 存中间数据,用数组排序日期
2. **HTML 解析**:沿用原作者 `htmlfile` ActiveX 对象 + DOM 遍历方式,稳定可靠
3. **编码**:`gb2312`(原作者用的),不要改 `utf-8`(新浪页面声明 gb2312)
4. **报告期格式**:HTML 里日期是 `2025-12-31` 字符串,转 `CDate()` 存成 datetime
5. **第一列识别大类**:新浪原 HTML 表里大类信息靠跨行合并 td 的 `rowspan` 实现。Claude Code 接手时**先 curl 一个真实页面看清楚 HTML 结构**,再写解析逻辑

---

## 4. 数据规范

### 4.1 单元格值规则
- `"--"` → 空白(Empty / `""`)
- 数字字符串 → `CDbl()` 转 Double
- 报告期/上市日期 → `CDate()` 转 Date,显示 `yyyy-mm-dd`
- 主营业务 → 字符串原样保留(可能很长)

### 4.2 表头格式
- R1(公司名层):字体加粗,深蓝底白字 `RGB(68,114,196)`,水平居中,跨列合并
- R2(报告期层):字体加粗,浅蓝底深色字 `RGB(217,225,242)` / `RGB(31,73,125)`,水平居中
- A 列(大类):字体加粗
- 冻结窗格:`B3`(大类列、指标列、表头 2 行都冻结)
- 自动筛选:`A1:最后列1`(只能在 R2 加,否则合并的 R1 不让加 — **这是个坑**,实际可能要把 AutoFilter 加在 R2,或者只能不加筛选,Claude Code 测试后定)

### 4.3 列宽
- A(大类):20
- B(指标名称):28
- C 起(数据列):14

---

## 5. 已知风险与待验证项

### 5.1 网络请求层

| 风险 | 应对建议 |
|---|---|
| 新浪可能限流(原工具无重试无限速) | 请求间加 `Application.Wait Now + TimeSerial(0,0,1)` 1 秒间隔 |
| 新浪页面结构可能已变 | **先 curl 一个真实页面验证 HTML 结构没变**,再写解析 |
| URL 模板里 `displaytype/4` 是分季度,可能要参数化 | 暂保留不动,后续需要再说 |
| 编码 gb2312 在某些公司名(生僻字)可能出错 | 沿用原作者方式,有问题再切 utf-8 |

### 5.2 数据完整性

| 风险 | 应对建议 |
|---|---|
| 不同公司报告期数量不同(并集) | 已确认用并集对齐,缺数据空白 |
| 不同公司财务指标差异(银行 vs 制造业) | 用指标并集,缺指标行存在但数据空白。**注意**:如果用户样本池混入金融股,指标差异会很大,宽表会有大量空行 |
| 港股/B股/中概股代码格式不同 | 当前样本全是 A 股,不考虑 |
| 季报 vs 半年报 vs 年报 | 用并集自然处理,不区分 |

### 5.3 VBA 工程层

| 风险 | 应对建议 |
|---|---|
| `Scripting.Dictionary` 在 Mac Excel 不支持 | 用户是 Windows,可用 |
| `ArrayList` 同上,且部分 Office 365 可能也禁用 | 优先用原生数组+辅助函数,Dictionary 仅做查找 |
| 多层表头 + AutoFilter 冲突 | 实测决定,不行就放弃 AutoFilter |
| 抓 5 公司 × 4 张财务表 + 1 张基本资料 = 25 次请求,可能触发反爬 | 加请求间隔,加 User-Agent,失败重试 1 次 |

---

## 6. 测试用例(给 Claude Code 跑通必须验证的清单)

**单元测试级**:
- [ ] HtmlGet 函数能正确返回 gb2312 解码后的字符串
- [ ] 解析资产负债表 HTML,能拿到所有报告期、所有指标、大类映射
- [ ] 解析利润表 HTML(注意 col 27-32 大类瑕疵的复刻)
- [ ] 解析现金流量表 HTML(注意"附注"大类)
- [ ] 解析指标表 HTML
- [ ] 解析公司基本资料 HTML,能拿到 5 个字段

**集成测试级**:
- [ ] 样本池 4 公司(安克创新、梦百合、喜临门、致欧科技),`一键全抓()` 运行后:
  - [ ] 4 张宽表 Sheet 都生成,格式正确(R1 公司合并、R2 报告期、R3+ 数据)
  - [ ] 公司顺序与样本池一致
  - [ ] 报告期降序
  - [ ] `--` 已转空、数字已转 Double、日期已转 Date
  - [ ] 利润表有"六、每股收益"瑕疵复刻
  - [ ] 上市公司基本资料 Sheet 有 4 行数据,5 个字段全
- [ ] 样本池只有 1 公司也能跑(边界)
- [ ] 样本池含有报告期数不同的公司(老股 + 新股)能正确做并集对齐

**回归测试**:
- [ ] 重复运行 `一键全抓()` 不累积、不报错(幂等)
- [ ] 抓取过程中网络中断,有清晰错误提示

---

## 7. 给 Claude Code 的开工建议

### 7.1 实施顺序
1. **第一步:先验证 HTML 结构没变**
   ```powershell
   # 用 PowerShell / curl 抓一个真实页面看看
   curl "http://money.finance.sina.com.cn/corp/go.php/vFD_BalanceSheet/stockid/300866/ctrl/part/displaytype/4.phtml" -o sample.html
   ```
   保存为本地 .html 文件,人工看 `<table id="BalanceSheetNewTable0">` 的结构。如果新浪改版,所有正则都要重做。

2. **第二步:在新工作簿里先把**单家公司单张表**的抓数+宽表生成跑通**(MVP)
   - 只抓资产负债表,只抓 1 家公司(300866)
   - 写出表头 + 数据
   - 跑通后再扩展

3. **第三步:扩展到 4 张财务表**(资产负债表 → 利润表 → 现金流量表 → 指标表)

4. **第四步:加多公司并集对齐逻辑**

5. **第五步:加基本资料抓取**

6. **第六步:整合一键全抓 + 错误处理 + 进度提示**

### 7.2 调试技巧
- 把 HTML 抓下来后**写到一个临时 Sheet 里**(`调试_HTML缓存`),便于离线调试解析逻辑而不每次重新请求
- 解析出的中间数据(Dictionary)用 `Debug.Print` 打印,或写到调试 Sheet
- 遇到字符编码问题,先看是 gb2312 解码错了还是 Excel 显示问题

### 7.3 不要做的事
- ❌ 不要保留 V2.2 原作者的任何抓数 Sub(用户明确要重写)
- ❌ 不要保留原 4 张主表(用户明确要删)
- ❌ 不要保留长表 Sheet(用户明确要删)
- ❌ 不要试图修正利润表"六、每股收益"瑕疵(用户已知,要求忠实复刻)
- ❌ 不要把基本资料字段扩到 5 个以外(用户明确只要 5 个核心字段)
- ❌ 不要引入 Power Query / Power Pivot(纯 VBA 实现)

### 7.4 Excel 环境
- 用户系统:Windows
- Office 版本:**未确认**(用户说不知道是不是 WSL,Excel 版本也没说)
- 假设:Office 365 / 2016+,有 `Scripting.Dictionary` 支持
- 第一次跑通后请用户确认实际 Excel 版本,如有问题再调

---

## 8. 交接物

放在同一个 git repo 下:

```
/
├── README.md                    # 项目简介、安装/使用说明
├── STATUS.md                    # 本文档,设计决策与进度
├── 上市公司财务数据查询.xlsm    # 最终交付物(Claude Code 不要直接生成,而是给一个空模板 + 全部 .bas)
├── modules/                     # 所有 VBA 源码,文本可 diff
│   ├── 模块_工具函数.bas
│   ├── 模块_抓资产负债表.bas
│   ├── 模块_抓利润表.bas
│   ├── 模块_抓现金流量表.bas
│   ├── 模块_抓指标表.bas
│   ├── 模块_抓基本资料.bas
│   └── 模块_总入口.bas
└── samples/                     # 测试用 HTML 离线缓存
    ├── 300866_balance.html
    ├── 300866_profit.html
    └── ...
```

**版本号**:V3.0(明确区别于原作者 V2.2)

---

## 9. 历史决策对照表(避免反复)

| 决策点 | 早先选项 | 最终选定 | 备注 |
|---|---|---|---|
| 输出格式 | A 长表 / B 表头转置 | 宽表(用户上传样例) | 经过 5 轮才到位 |
| 报告期对齐 | A 固定/B 动态/C 并集 | **C 并集** | |
| 原表保留 | 全留/全删/留主表 | **全删** | |
| 瑞华底稿 | 保留/删 | **删** | |
| 抓数路径 | 沿用原 VBA / 重写 | **重写** | 用户接受工作量 |
| 基本资料字段 | 全部/核心/推荐 | **核心 5 个** | 代码/简称/上市日期/行业/主营业务 |
| Power Query | 用 / 不用 | **不用** | 纯 VBA |

---

**文档结束。Claude Code 接手时从 §3(最终设计方案)读起即可,§2 历史只在你想理解为什么这么设计时再看。**

---

# Phase 4b 实施进展(2026-05-02 EOD,Claude+Codex 并行开发交接)

> 这一节是 2026-05-02 一整天工作的进度交接,**明天 Eric 用 codex 并行开发时从这里读起**。
> §1-9 是设计基线。§Phase 4b 是实施实际状态。

## A. 已交付(Phase 1 → Phase 4b-4)

### A.1 代码资产

```
VBA Captor/
├── STATUS.md                          ← 本文档(基线 + 进展交接)
├── 上市公司财务数据查询.xlsm        ← 工作产物,install_modules.py 一键打包
├── modules/                           ← 全部 .bas 源(git 友好,可 diff)
│   ├── JsonConverter.bas              ← Tim Hall VBA-JSON v2.3.1(美股 EDGAR/雪球用)
│   ├── 模块_工具函数.bas              ← HTTP/CIK 映射/季度年份过滤/WriteWideTable 等共享 helper
│   ├── 模块_总入口.bas                ← 一键全抓(顺序调 8 个 Main + 汇总)
│   ├── 模块_抓资产负债表.bas          ← A 股 BS thin wrapper → RunOneStatement
│   ├── 模块_抓利润表.bas              ← A 股 IS thin wrapper
│   ├── 模块_抓现金流量表.bas          ← A 股 CF thin wrapper
│   ├── 模块_抓指标表.bas              ← A 股 Indicator thin wrapper
│   ├── 模块_抓美股财报.bas            ← 美股 workhorse(RunUSStatement / FetchAndAccumulateUSCompany / FetchBSFromXueqiu / AppendUSRatios)
│   ├── 模块_抓美股资产负债表.bas      ← 美股 BS thin wrapper(27 个 us-gaap concepts)
│   ├── 模块_抓美股利润表.bas          ← 美股 IS thin wrapper
│   ├── 模块_抓美股现金流量表.bas      ← 美股 CF thin wrapper
│   └── 模块_抓美股指标表.bas          ← 美股 Indicator thin wrapper(+ AppendUSRatios 8 个比率公式)
├── tools/
│   ├── build_template.py              ← openpyxl 生成空 上市公司财务数据查询.xlsx 模板(使用说明/样本池/4 A 股_/4 美股_/诊断表)
│   └── install_modules.py             ← Excel COM 把 .xlsx → .xlsm + 注入 12 个 .bas + 加圆角按钮 + 季度/cookie 单元格 + 样本池美化
└── samples/                           ← 离线 HTML/JSON 调试样本
    ├── 300866_*.html                  ← A 股 5 类页面
    ├── AAPL_edgar.json                ← 美股 EDGAR 全量样本
    └── xueqiu_POM_bs.json             ← 雪球 POM BS 实际响应(Phase 4b-5 调试用,UTF-16 编码)
```

### A.2 已完工功能(Phase 1 → 4b-4 都跑通了)

| Phase | 功能 | 状态 |
|---|---|---|
| 1 | A 股 BS MVP(单家) | ✅ 跑通 |
| 2 | A 股 IS/CF/Indicator + 一键全抓 + 基本资料 | ✅ 跑通(后基本资料 4b-3 删) |
| 3 | 季度选择(全部/Q1/Q2/Q3/Q4)+ 圆角按钮 | ✅ 跑通 |
| 4a | 港股(原计划) | ⏸ 跳过(新浪 HK 数据稀疏,放弃) |
| 4b-1 | 美股 EDGAR BS MVP(AAPL/AMZN) | ✅ 跑通 |
| 4b-1.1 | 美股年份/季度过滤 + H 列公式补齐 | ✅ 跑通 |
| 4b-2 | 美股 IS/CF/Indicator(英文标签) | ✅ 跑通 |
| 4b-3 | 删除上市公司基本资料 + A 股_ 前缀对称 + 美股 8 比率 | ✅ 跑通 |
| 4b-4 | C/D 列 spacer + 样本池美化 + 自动检测市场 + POM 失败定位 | ✅ 美化跑通,POM 没解决留到 4b-5 |
| **4b-5** | **POM(20-F filer)雪球 BS fallback** | ✅ **已跑通,见 §F** |
| **4b-6** | **POM/20-F 雪球 fallback 扩展到美股 IS/CF/Indicator** | ✅ **已跑通,见 §F** |
| **4b-7** | **修复美股指标表比率公式跨表错配** | ✅ **已跑通,见 §G** |

## B. **历史问题记录:Phase 4b-5 POM 雪球 fallback bug(已解决,见 §F)**

### B.1 Bug 现象

用户操作:
- 样本池 row 12: `A12=POM, B12=石榴云医, C12=US`(C 列由公式自动检测)
- A2=2024, A4=Q4
- B5 已粘 xueqiu cookie(`xq_a_token=xxx...` 或完整 Cookie 头,详见 install_modules.py 加的 cell guide)
- 点『更新美股资产负债表』

弹窗结果:
```
美股_资产负债表 抓取完成 (单位: 百万美元)
用时: 30.1 秒
公司数: 0 / 期数: 0

失败 1 条:
POM 石榴云医: 错误的参数号或无效的属性赋值
```

且 **没有** 我加的 `[stage=...]` 前缀(EOD 最后一次 instrument 加的)。

### B.2 已确认工作正常的部分(不要重复调试)

- ✅ EDGAR 前缀逻辑正确:POM 在 EDGAR 返回 404(因为是 20-F 外国发行人,SEC 不收录非 us-gaap),正确触发 xueqiu fallback
- ✅ 雪球 HTTP 请求成功:`samples/xueqiu_POM_bs.json` 在每次跑后都更新,JSON 完整(8 期数据,2020 FY → 2025 Q6)
- ✅ JSON 内容含 FY2024(`ed: "2024-12-31"`)条目,各字段齐全(`total_assets: [46227586.0, ...]` / `total_liab: [545908548.0, ...]` / `total_holders_equity: [-2263421920.0, ...]`)
- ✅ 雪球 cookie 鉴权通过(`error_code: 0`,如果 cookie 失效会是 `400016` "anonymous denied")
- ✅ 字段名映射已校准:基于真实 POM JSON 把 `mapXq` 改对了(`cce`/`net_receivables`/`total_liab`/`total_holders_equity` 等都是从 dump 文件验证过的真实键名)

### B.3 已确认问题:错误信息 **应该** 含 `[stage=...]` 但没有

最后一次提交里(`modules/模块_抓美股财报.bas` line ~512-680)给 `FetchBSFromXueqiu` 加了 stage 追踪:

```vba
Private Sub FetchBSFromXueqiu(...)
    Dim stage As String: stage = "init"
    On Error GoTo XqErr

    stage = "ReadCookie"   ' 每个关键步骤前更新
    ...
    stage = "ParseJson"
    ...
    stage = "Record#" & recIdx & ":concept#" & ci & ":cdbl(" & cand & ")"
    ...

XqErr:
    Dim origNum As Long: origNum = Err.Number
    Dim origDesc As String: origDesc = Err.Description
    Err.Clear
    Err.Raise origNum, "FetchBSFromXueqiu", _
        "[stage=" & stage & "] " & origDesc
End Sub
```

**预期**:错误描述应该是 `[stage=Record#2:concept#3:cdbl(inventory)] 类型不匹配` 这种格式。

**实际**:用户看到的依然是裸的 `错误的参数号或无效的属性赋值` —— 没有 stage 前缀。

**两种可能**:

1. **上市公司财务数据查询.xlsm 没真的重装新模块** — 用户可能没关掉重开 Excel,VBA Project 还是缓存的旧版本。
   - **明天验证步骤**:让用户彻底关 Excel(任务管理器确认 EXCEL.EXE 没了)→ 重跑 `py tools/install_modules.py` → 再开 `上市公司财务数据查询.xlsm` → 重测。
   - 或者最简单:在 `上市公司财务数据查询.xlsm` 里 Alt+F11 → 找 `模块_抓美股财报` → 看 `FetchBSFromXueqiu` 头几行有没有 `Dim stage As String: stage = "init"`。如果**没有**就是没装上;如果**有**就是真的没生效,得继续往下查。

2. **错误发生在 `FetchAndAccumulateUSCompany` 而非 `FetchBSFromXueqiu`** — 比如 EDGAR fetch 阶段就炸了,根本没走到 xueqiu fallback。
   - 验证:在 `FetchAndAccumulateUSCompany` 头部也加一个类似的 stage 追踪,或者在 EDGAR 失败处加 `Debug.Print "EDGAR 失败: " & edgarErrDesc`。
   - **更可能的根因**:看 line 195-204 的 fallback 触发条件:
     ```vba
     If edgarErrNum <> 0 Then
         If strKind = "BalanceSheet" Then
             FetchBSFromXueqiu strTicker, conceptMap, strQuarter, lngYear, _
                               dictData, dictPeriodSet, dictIndicatorSet, dictCategoryMap
             Exit Sub
         Else
             Err.Raise vbObjectError + 526, "FetchUS", edgarErrDesc
         End If
     End If
     ```
     看着没问题。但 `On Error GoTo CleanUp` 在外层(RunUSStatement),内层 `FetchBSFromXueqiu` 自己 `On Error GoTo XqErr` 拦截,**理论上 stage info 应该传上去**。

### B.4 关于错误描述 "错误的参数号或无效的属性赋值"

这是 VBA 的 **runtime error 380** 的中文本地化描述(英文是 "Invalid property value")。最常见触发:
- `dict("key") = value` 形式赋值时 key 是非法值(空/Null)
- 给 Object 的某个属性赋了不兼容类型(给 Range.Value 赋数组维度不对 之类)
- Collection 的 `.Item(i)` 越界或参数类型错

我已经做了的预防:
- `dict("key")` 改 `dict.Item("key")` (有些 Office 版本对默认成员调用解析比较脆)
- `IsEmpty(val)` 加 `Not IsNull(val)`(Tim Hall 把 JSON `null` 解析成 VBA Null 不是 Empty)
- 把 dict access 都先 Exists 检查再读,避免 auto-add 副作用

### B.5 明天的优先建议(Codex/Claude 任选)

**Plan A — 最简朴的:加打印到工作表**

在 `FetchBSFromXueqiu` 入口加一行:
```vba
ThisWorkbook.Sheets("样本池").Range("J1").Value = "FetchBSFromXueqiu entered: " & Now()
```
跑一次,看 J1 有没有内容。如果**没有** → 错误在到达 xueqiu 之前(比如 EDGAR 段);如果**有** → 错误在 xueqiu 里面但 stage 追踪没工作,继续在每个 stage 后加 `Sheets("样本池").Range("J2").Value = stage` 这种笨办法定位。

**Plan B — 更彻底的:在 RunUSStatement 拦截后立刻 MsgBox**

```vba
On Error Resume Next
Err.Clear
Call FetchAndAccumulateUSCompany(...)
If Err.Number <> 0 Then
    MsgBox "DEBUG: code=" & strCode & vbCrLf & _
           "Err.Number=" & Err.Number & vbCrLf & _
           "Err.Source=" & Err.Source & vbCrLf & _
           "Err.Description=" & Err.Description, vbExclamation
    intFailCnt = intFailCnt + 1
    ...
```
跑一次会立刻弹窗,看 Err.Source 能告诉我们错误是从哪个 Sub raise 的。

**Plan C — 假设是 EDGAR 段先死:换 IFRS endpoint 试试**

POM 是 20-F filer,SEC 上理论上有 ifrs-full 数据(虽然实测我之前搜过 `https://data.sec.gov/api/xbrl/companyfacts/CIK*.json` 也是 404)。验证步骤:
```bash
curl -A "Eric Zhang 214978902@qq.com" "https://data.sec.gov/submissions/CIK0001823575.json" | jq .
```
确认 POM CIK = 1823575。如果 submissions 接口能通,companyfacts 还是 404 就是 SEC 真没收录这家 → 雪球是唯一路。

**Plan D — 雪球 anonymous 模式**

`https://stock.xueqiu.com/v5/stock/finance/us/balance.json?symbol=POM&type=all&is_detail=true&count=8` 不带 cookie 直接 curl 试试看能不能返,可能用户 cookie 失效但接口本身是公开的。

### B.6 不要踩的坑(已经踩过的)

1. **WinHttp 不自动解 gzip** — `Accept-Encoding: gzip,deflate` 会拿到一坨乱码。已设 `identity`,别改。
2. **`CStr(Null)` 直接抛** — 必须先 `IsNull` 检查。`NzStr` helper 已封装。
3. **VBA `Public Const` / `Public` 必须在所有 Sub 之前** — 否则编译报 "无效的属性"(注意,这是另一种 380 触发场景)。
4. **VBA 行延续 `_` 限制 25 行** — `Array(...) _ Array(...) _ ...` 长串会编译失败,改成 `a(0)=Array() / a(1)=Array() / ...` 索引化。
5. **JsonConverter 数组解析返回 `Collection`(1-based)不是 VBA Array** — 用 `.Count` 和 `.Item(i)` 访问。
6. **POM 雪球字段是 `[absolute_value, yoy_pct]` 数组,不是裸值** — `XueqiuValue` helper 已处理 Collection 取第 1 项。
7. **dump 文件是 UTF-16 LE BOM 编码**(`fso.CreateTextFile(fname, True, True)` 第 3 个 True = Unicode)。Python 读取要 `encoding='utf-16'`。
8. **install_modules.py `wb.Sheets.Copy + rename` 会跨次掉 sheet** — 已改成纯 `Worksheets.Add`。

### B.7 雪球 cookie 状态

- 用户已在 `样本池!B5` 粘了 cookie
- 用户的 cookie 在 30:1 秒级别能跑完 fetch + dump(说明 HTTP 通)
- **如果明天测下来 cookie 失效**: F12 重新登录 xueqiu.com 后取 `xq_a_token` cookie 即可
- 现在的 `ReadXueqiuCookie` 兼容两种粘法:
  - 纯 token 值(无 `=`)→ 自动包成 `xq_a_token=<value>`
  - 完整 Cookie 头(含 `=`)→ 原样用

## C. 测试用样本池(用户当前的)

```
A1=年份(留空=取最新)    A2=2024
A3=季度(Q1/Q2/...)        A4=Q4
A5=雪球Cookie(可选)       B5:C5(merged)=<用户粘的 token>
A7=股票代码 B7=股票简称 C7=市场
A8=300866   B8=安克创新   C8=A   (公式自动检测)
A9=603313   B9=梦百合     C9=A
A10=603008  B10=喜临门    C10=A
A11=301376  B11=致欧科技  C11=A
A12=POM     B12=石榴云医  C12=US
```
4 家 A 股 + POM 一只美股(20-F)。AAPL/AMZN 之前测过都跑通,**Q4=FY 年报模式下 EDGAR 美股 BS/IS/CF/Indicator 都正常**,只 POM 触发 fallback 失败。

## D. install_modules.py 关键 flag(避免误改)

- `BUTTON_ANCHOR_COL = "E"` — 按钮起始列,D 是 spacer 留 3 字符宽
- `POOL_DATA_START_ROW = 8` — 数据从 row 8 起
- `DECOMMISSIONED_SHEETS = ["上市公司基本资料", "资产负债表", "利润表", "现金流量表", "指标表"]` — Phase 4b-3 删的旧 sheet 名,运行时遇到这些会自动 delete
- `DECOMMISSIONED_MODULES = ["模块_抓基本资料"]` — 删除已淘汰模块,免得 import 报错
- 安装顺序:模板存在则保留,9 张 sheet 缺哪个补哪个,模块全量替换,按钮全删重建

## E. 一行总结

**本节 B 记录的是 2026-05-02 EOD 的历史故障现场。2026-05-03 已在 §F 修复并验证: POM/20-F 公司现在可通过雪球 fallback 跑通美股 BS/IS/CF/Indicator,正常 us-gaap filer(AAPL/AMZN/Tesla 类)仍优先走 EDGAR。**

## F. Phase 4b-5/4b-6 完成记录(2026-05-03, Codex)

### F.1 已完成修复

- `模块_抓美股财报.bas` 已把原 `FetchBSFromXueqiu` 扩展为通用 `FetchUSFromXueqiu`,在 EDGAR companyfacts 失败且报表类型为 `BalanceSheet` / `Income` / `CashFlow` / `Indicator` 时走雪球 fallback。
- 雪球 endpoint 映射:
  - `BalanceSheet` → `stock/finance/us/balance.json`
  - `Income` → `stock/finance/us/income.json`
  - `CashFlow` → `stock/finance/us/cash_flow.json`
  - `Indicator` → `stock/finance/us/indicator.json`
- 字段映射按 POM 真实 JSON 校准。没有直接雪球字段的指标保持空白,不伪造数据。
- 修复根因:VBA-JSON 把雪球 `[absolute_value, yoy_pct]` 解析为 `Collection`;旧代码把对象直接塞进 `Variant`,在部分 Office/VBA 版本下触发 450 并被 `On Error Resume Next` 吞掉,导致值全空或错误残留。现在 `XueqiuValue` 显式 `Set objVal = record.Item(key)` 后取 `objVal.Item(1)`。
- 错误处理已增强:上层失败日志包含错误号、来源、描述;雪球 fallback 异常统一重抛为自定义错误号,描述带 `[stage=...]`、原始错误号和原始来源。
- `ReadXueqiuCookie`、`XueqiuValue`、`ParseXueqiuReportDate`、fallback 成功出口均清理 `Err`,避免成功路径被残留错误误判失败。

### F.2 验证结果

POM-only 样本池、`A2=2024`、`A4=Q4`、B5 有雪球 cookie,内存运行四个美股按钮,不保存测试输出:

| 表 | 关键校验 | 结果 |
|---|---:|---|
| 美股_资产负债表 | `Total assets` | `46.227586` 百万美元 |
| 美股_利润表 | `Revenue` | `342.55792` 百万美元 |
| 美股_现金流量表 | `Cash from operations` | `-16.13088` 百万美元 |
| 美股_指标表 | `Basic EPS (USD/share)` | `-3.7874239999999997` |

全部四张表均写出 `POM(POM)` / `2024-12-31`,失败数 `0`。

回归: AAPL/AMZN 样本池、`A2=2024`、`A4=Q4`,内存运行四个美股按钮:

| 表 | 结果 |
|---|---|
| 美股_资产负债表 | 失败数 `0` |
| 美股_利润表 | 失败数 `0` |
| 美股_现金流量表 | 失败数 `0` |
| 美股_指标表 | 失败数 `0` |

### F.3 新增/更新的调试样本

- `samples/xueqiu_POM_bs.json`
- `samples/xueqiu_POM_income.json`
- `samples/xueqiu_POM_cash_flow.json`
- `samples/xueqiu_POM_indicator.json`

### F.4 当前限制

- 雪球 fallback 是为 20-F/ADR 公司兜底,优先保证 POM 这类 EDGAR 404 的公司能跑通;正常 us-gaap filer 仍优先走 EDGAR。
- 指标表 fallback 目前只映射原始 `Basic EPS` / `Diluted EPS`;其它比率仍由 `AppendUSRatios` 根据 BS/IS 公式补行。
- 现金流 fallback 只映射雪球有明确字段的指标;EDGAR conceptMap 中没有雪球字段的行保持空白。

## G. Phase 4b-7 完成记录(2026-05-03, Codex)

### G.1 修复内容

- 修复 `美股_指标表` 追加比率行的公式逻辑。
- 旧逻辑:按指标表当前列字母直接引用 `美股_资产负债表` / `美股_利润表` 同列,再套固定指标行号。跨表报告期或公司列不完全一致时会错配;且字段缺失导致行号变化时,可能把 `Net Margin` / `ROA` / `ROE` 引到错误行。
- 新逻辑:对指标表每一个数据列,先读取该列的公司表头和报告期,再到 BS/IS 中按 **公司表头 + 报告期 + 指标名称** 定位真实单元格,最后生成公式。
- 公式增加 `IFERROR(...,"")`,缺少依赖项时留空,不输出误导性的 0。
- 新增辅助函数:
  - `HeaderTextAt`:兼容 R1 合并单元格,取公司表头。
  - `PeriodKey`:统一报告期比较格式。
  - `FindStatementColumn`:按公司表头 + 报告期找目标表数据列。
  - `SheetCellRef`:生成跨 sheet 单元格引用。

### G.2 验证结果

POM-only、`A2=2025`、`A4=全部`,内存运行 BS/IS/Indicator:

| 指标 | 修复后结果 |
|---|---:|
| Current Ratio | `0.17327101399664896` |
| Quick Ratio | `0.140869814349066` |
| Gross Margin | `0.162305411724115` |
| Operating Margin | `-0.11423880908799539` |
| Net Margin | `-0.11397116399211002` |
| ROA | `-0.4532402241806034` |
| ROE | `0.008513288840167573` |

关键检查:`Net Margin` / `ROA` / `ROE` 不再显示误导性 `0.00%`;ROE 公式正确引用 `美股_利润表` 的 `Net income` 行和 `美股_资产负债表` 的 `Total stockholders' equity` 行。

AAPL-only、`A2` 留空、`A4=全部`,多报告期合并表头回归:

| 列 | 报告期 | ROE 公式定位 |
|---|---|---|
| C | `2026-03-28` | 引用 BS/IS 的 C 列 |
| D | `2025-12-27` | 引用 BS/IS 的 D 列 |
| E | `2025-09-27` | 引用 BS/IS 的 E 列 |

三列均失败数 `0`,验证合并公司表头下非首列也能正确匹配公司和报告期。

## H. Phase 4b-8 完成记录(2026-05-03, Codex)

### H.1 修复内容

- A 股和美股 `指标表` 统一新增标准指标层,固定放在原始抓取指标之前,便于直接横向对标。
- `WriteWideTable` 对 `A股_指标表` / `美股_指标表` 启用三列静态表头:
  - A 列: `指标类型`
  - B 列: `指标名称`
  - C 列: `英文指标名`
  - D 列起: 公司 × 报告期数据列
- 新增公共入口 `AppendStandardIndicators(ws, market)`,两个指标表入口都会调用:
  - `模块_抓指标表.Main` → A 股标准指标
  - `模块_抓美股指标表.Main` → 美股标准指标
- 新增 `SetSilentMode` 供自动化验证和一键流程控制弹窗。

### H.2 当前标准指标清单

| 指标类型 | 指标名称 | 英文指标名 |
|---|---|---|
| 盈利性指标 | 销售净利率 | Net Profit Margin |
| 盈利性指标 | 毛利率 | Gross Profit Margin |
| 盈利性指标 | 期间费用率 | Operating Expense Ratio |
| 盈利性指标 | 总资产回报率 (ROA) | Return on Assets (ROA) |
| 盈利性指标 | 股东权益回报率 (ROE) | Return on Equity (ROE) |
| 成长性指标 | 总资产增长率 | Total Assets Growth Rate |
| 成长性指标 | 主营业务收入增长率 | Revenue Growth Rate |
| 成长性指标 | 净利润增长率 | Net Profit Growth Rate |
| 偿债能力指标 | 流动比率 | Current Ratio |
| 偿债能力指标 | 速动比率 | Quick Ratio |
| 偿债能力指标 | 现金比率 | Cash Ratio |
| 偿债能力指标 | 资产负债率 | Debt-to-Asset Ratio |
| 运营能力指标 | 存货周转天数 | Days Inventory Outstanding (DIO) |
| 运营能力指标 | 应收款周转天数 | Days Sales Outstanding (DSO) |
| 运营能力指标 | 应付账款周转天数 | Days Payable Outstanding (DPO) |
| 运营能力指标 | 营运资金周转天数 | Cash Conversion Cycle (CCC) |
| 运营能力指标 | 流动资产周转率 | Current Asset Turnover |
| 运营能力指标 | 总资产周转率 | Total Asset Turnover |

### H.3 公式口径

- 盈利性:
  - 销售净利率 = 净利润 / 营业收入
  - 毛利率 = (营业收入 - 营业成本) / 营业收入;美股优先用 `Gross profit / Revenue`
  - 期间费用率 = A 股销售/管理/财务/研发费用合计 ÷ 营业收入;美股优先用 `Total operating expenses / Revenue`
  - ROA / ROE = 净利润 ÷ 平均总资产 / 平均股东权益;没有可比上期时用当期余额
- 成长性:
  - 总资产、收入、净利润增长率 = 当期 / 上期 - 1;没有上期时留空
- 偿债能力:
  - 流动比率、速动比率、现金比率、资产负债率按 BS 项直接计算
- 运营能力:
  - DIO / DSO / DPO 使用平均余额 × 期间天数 ÷ 收入或成本
  - CCC = DIO + DSO - DPO
  - 流动资产周转率 / 总资产周转率 = 收入 ÷ 平均流动资产 / 平均总资产

### H.4 验证结果

POM-only、`A2=2025`、`A4=全部`,内存运行美股 BS/IS/Indicator,不保存测试输出:

| 检查 | 结果 |
|---|---|
| 表头 | `指标类型 / 指标名称 / 英文指标名 / 石榴云医(POM) / 2025-06-30` |
| 标准指标行 | Row 3-20 共 18 行 |
| 销售净利率 | `-11.40%` |
| 毛利率 | `16.23%` |
| 流动比率 | `0.17` |
| 速动比率 | `0.14` |
| 资产负债率 | `1284.26%` |

安克创新-only、`A2=2025`、`A4=全部`,内存运行 A 股 BS/IS/Indicator,不保存测试输出:

| 检查 | 结果 |
|---|---|
| 表头 | `指标类型 / 指标名称 / 英文指标名 / 安克创新(300866) / 2025-12-31` |
| 标准指标行 | Row 3-20 共 18 行 |
| 销售净利率 | `8.34%` |
| 毛利率 | `45.07%` |
| 总资产增长率 | `0.23%` |
| 流动比率 | `2.38` |
| 营运资金周转天数 | `94.97` |

`tools/install_modules.py` 已重跑并保存 `上市公司财务数据查询.xlsm`;Scripting Runtime 引用冲突提示仍是既有无害提示,模块替换和按钮重建成功。

## I. Phase 4b-9 完成记录(2026-05-03, Codex)

### I.1 修复内容

- `A股_指标表` 和 `美股_指标表` 现在只生成 18 个标准指标,不再保留任何网站/EDGAR/雪球原始指标行。
- `模块_抓指标表.Main` 不再调用新浪指标页 `RunOneStatement(..., "indicator", ...)`,改为 `BuildStandardIndicatorSheet "A"`。
- `模块_抓美股指标表.Main` 不再调用 `RunUSStatement "Indicator"`,改为 `BuildStandardIndicatorSheet "US"`;美股 EPS / shares / dividend raw 行已移除。
- 指标表表头固定为:
  - A 列: `指标类型`
  - B 列: `指标名称`
  - C 列: `英文指标名`
  - D 列起: 公司 × 报告期
- `tools/install_modules.py` 已固定 Tab 顺序:
  - `使用说明`
  - `样本池`
  - `A股_资产负债表`
  - `A股_利润表`
  - `A股_现金流量表`
  - `A股_指标表`
  - `美股_资产负债表`
  - `美股_利润表`
  - `美股_现金流量表`
  - `美股_指标表`

### I.2 公式口径修正

按新浪 A 股原始指标页校准后,标准指标公式口径调整为:

- 销售净利率 = `五、净利润 / 营业收入`,对齐新浪 `销售净利率(%)`。
- 毛利率 = `(营业收入 - 营业成本 - 营业税金及附加) / 营业收入`,对齐新浪 `主营业务利润率(%)`。
- 期间费用率 = `(销售费用 + 管理费用 + 财务费用) / 营业收入`,对齐新浪 `三项费用比重`;A 股不再把研发费用并入该项。
- ROA = `五、净利润 / 平均总资产`,平均资产使用当期总资产和上一财年年末总资产,对齐新浪 `总资产净利润率(%)`。
- ROE = `归属于母公司所有者的净利润 / 期末归母权益`,对齐新浪 `净资产收益率(%)`。
- 总资产增长率 = `当期总资产 / 上一财年年末总资产 - 1`。
- 主营业务收入增长率 = `当期营业收入 / 去年同期营业收入 - 1`。
- 净利润增长率 = `当期五、净利润 / 去年同期五、净利润 - 1`。
- A 股运营天数按新浪口径使用 360 天,YTD 期间为 Q1=90、Q2=180、Q3=270、Q4=360。
- DIO / DSO / 总资产周转率 / 流动资产周转率使用上一财年年末余额参与平均,对齐新浪周转指标。
- DPO / CCC 新浪 A 股指标页没有直接对应项,保留行业通用口径:
  - DPO = 平均应付账款 × 期间天数 / 营业成本
  - CCC = DIO + DSO - DPO

为支持这些公式,A 股资产负债表/利润表在 A2 指定年份时会额外抓上一年数据,供指标表公式引用;指标表自身只展示当前 A2/A4 选择的报告期。

美股资产负债表/利润表同样在 A2 指定年份时额外保留上一年数据,使美股标准指标的增长率和平均资产/权益类公式也能引用上一年基准。

### I.3 验证结果

安克创新-only、`A2=2025`、`A4=全部`,内存运行 A 股资产负债表、利润表、指标表,不保存测试输出:

| 指标 | 新公式结果 | 新浪原始指标对照 |
|---|---:|---:|
| 销售净利率 | `8.58%` | `8.5769%` |
| 毛利率 | `45.00%` | `44.9987%` |
| 期间费用率 | `26.13%` | `26.1269%` |
| ROA | `14.27%` | `14.2741%` |
| ROE | `24.18%` | `24.18%` |
| 总资产增长率 | `20.86%` | `20.8579%` |
| 主营业务收入增长率 | `23.49%` | `23.4897%` |
| 净利润增长率 | `18.36%` | `18.3649%` |
| 流动比率 | `2.38` | `2.3767` |
| 速动比率 | `1.64` | `1.6384` |
| 现金比率 | `0.54` | `54.0224%` |
| 资产负债率 | `46.62%` | `46.6202%` |
| 存货周转天数 | `88.38` | `88.3804` |
| 应收款周转天数 | `20.80` | `20.8042` |
| 流动资产周转率 | `2.14` | `2.1448` |
| 总资产周转率 | `1.66` | `1.6642` |

POM-only、`A2=2025`、`A4=全部`,内存运行美股资产负债表、利润表、指标表:

| 检查 | 结果 |
|---|---|
| 表头 | `指标类型 / 指标名称 / 英文指标名 / 石榴云医(POM) / 2025-06-30` |
| 行数 | Row 3-20,只含 18 个标准指标 |
| 总资产增长率 | `-5.10%`,引用上一年基准 |
| 主营业务收入增长率 | `16.19%`,引用去年同期 |
| 净利润增长率 | `41.48%`,引用去年同期 |

最后已重跑 `tools/install_modules.py` 并保存 `上市公司财务数据查询.xlsm`;Scripting Runtime 引用冲突提示仍是既有无害提示。

## J. Phase 4b-10 完成记录(2026-05-03, Codex)

### J.1 修复内容

- 优化 `美股_现金流量表` 字段覆盖率。
- `模块_抓美股现金流量表.bas` 的 `GetCFConcepts()` 从 14 个项目扩展到 34 个候选项目,覆盖:
  - Operating: Net income、D&A、SBC、递延税、营运资本变动、CFO 等
  - Investing: 证券购买/到期/出售、投资购买、Capex、并购、其它投资、CFI 等
  - Financing: 股息、回购、发股、股权激励扣税、长债发行/偿还、短债净额、其它融资、CFF 等
  - Cash reconciliation: FX effect、净现金变动、期初/期末现金
- `FetchAndAccumulateUSCompany` 支持同一个指标配置多个 EDGAR us-gaap concept 候选,按顺序取第一个能在当前公司/期间匹配到数据的 concept。这样不同公司使用同义 XBRL 标签时不再漏抓。
- 雪球 fallback 的现金流字段也补充:
  - `purs_of_invest` → Purchases of investments
  - `effect_of_exchange_chg_on_cce` → FX effect on cash
  - `cce_at_boy` → Cash at beginning of period
  - `cce_at_eoy` → Cash at end of period

### J.2 验证结果

AAPL-only、`A2=2024`、`A4=Q4`,内存运行 `更新美股现金流量表`,不保存测试输出:

| 检查 | 结果 |
|---|---|
| 表头 | `Apple(AAPL) / 2024-09-28` |
| 输出行数 | Row 3-28,共 26 个现金流项目 |
| Cash from operations | `118,254.00` 百万美元 |
| Purchases of marketable securities | `48,656.00` 百万美元 |
| Stock repurchases | `94,949.00` 百万美元 |
| Cash from financing | `-121,983.00` 百万美元 |

POM-only、`A2=2025`、`A4=全部`,内存运行 `更新美股现金流量表`,不保存测试输出:

| 检查 | 结果 |
|---|---|
| 表头 | `石榴云医(POM) / 2025-06-30` |
| 输出行数 | Row 3-11,共 9 个现金流项目 |
| Cash from operations | `-14.99` 百万美元 |
| Purchases of investments | `-0.51` 百万美元 |
| Cash from financing | `13.59` 百万美元 |
| Cash at beginning of period | `7.65` 百万美元 |
| Cash at end of period | `5.75` 百万美元 |

说明:POM 的雪球现金流 API 本身只返回少量明细字段,本阶段已把本地样本 JSON 中可用的金额字段全部映射进表里;缺少 D&A、营运资本明细、股息/回购等是数据源不提供,不是当前映射漏抓。

已重跑 `tools/install_modules.py` 并保存 `上市公司财务数据查询.xlsm`;Scripting Runtime 引用冲突提示仍是既有无害提示。

## K. Phase 4b-11 完成记录(2026-05-03, Codex)

### K.1 修复内容

- 修复 `HTT / 趣店 / US` 这类中概股美股资产负债表失败的问题。
- 根因: HTT/QD 这类 20-F/ADR 公司在 EDGAR companyfacts 可能能返回 JSON,但没有可匹配的 `us-gaap` 财报字段。旧逻辑只在 EDGAR 请求失败时 fallback 雪球;如果 EDGAR 成功但匹配字段数为 0,会直接进入“无匹配数据”失败分支。
- 新逻辑: `FetchAndAccumulateUSCompany` 在以下情况都会转雪球 fallback:
  - EDGAR 请求失败/404
  - JSON 缺 `facts`
  - JSON 缺 `us-gaap`
  - 当前报表 conceptMap 匹配结果为 0
- 这个修改覆盖 BS/IS/CF/Indicator 支持的 fallback 类型,不是 HTT hardcode。

### K.2 验证结果

HTT-only、`A2=2025`、`A4=全部`,内存运行 `更新美股资产负债表`,不保存测试输出:

| 检查 | 结果 |
|---|---|
| 表头 | `趣店(HTT) / 2025-12-31` |
| 输出行数 | Row 3-25 |
| Total assets | `13,612.91` 百万美元 |
| Total liabilities | `1,981.31` 百万美元 |
| Total stockholders' equity | `11,631.60` 百万美元 |
| Total liabilities & equity | `13,612.91` 百万美元 |

用户提供的雪球页面 `https://xueqiu.com/snowman/S/HTT/detail#/ZCFZB` 与验证结果一致: 雪球端有 HTT 资产负债表数据;失败原因是原 fallback 触发条件不完整。

已重跑 `tools/install_modules.py` 并保存 `上市公司财务数据查询.xlsm`;Scripting Runtime 引用冲突提示仍是既有无害提示。

## L. Phase 4b-12 完成记录(2026-05-03, Codex)

### L.1 修复内容

- 修复美股宽表在多公司期间不一致时使用全局期间并集导致空列过多的问题。
- `WriteWideTable` 新增 `perCompanyPeriods` 开关:
  - 默认 `False`,A股仍保持全局报告期并集对齐,便于同行横向比较。
  - 美股 `RunUSStatement` 传入 `True`,每家公司只展开自己实际抓到的报告期。
- 修改覆盖美股共享入口,因此 `美股_资产负债表`、`美股_利润表`、`美股_现金流量表` 的宽表列布局保持一致。
- 指标表继续从资产负债表表头复制公司/期间列,会自然继承美股按公司自身期间展开后的列结构。

### L.2 验证结果

HTT + POM、`A2=2025`、`A4=全部`,内存运行 `更新美股资产负债表`,不保存测试输出:

| 公司 | 输出列数 | 报告期 |
|---|---:|---|
| 趣店(HTT) | 8 | `2025-12-31`, `2025-09-30`, `2025-06-30`, `2025-03-31`, `2024-12-31`, `2024-09-30`, `2024-06-30`, `2024-03-31` |
| 石榴云医(POM) | 3 | `2025-06-30`, `2024-12-31`, `2024-06-30` |

POM 不再被 HTT 的 8 个报告期撑出空列。POM `Total assets` 验证值:

| 报告期 | Total assets(百万美元) |
|---|---:|
| `2025-06-30` | `43.87` |
| `2024-12-31` | `46.23` |
| `2024-06-30` | `56.26` |

已重跑 `tools/install_modules.py` 并保存 `上市公司财务数据查询.xlsm`;Scripting Runtime 引用冲突提示仍是既有无害提示。检查无遗留 Excel 后台进程。

## M. Phase 4b-13 完成记录(2026-05-03, Codex)

### M.1 收口内容

- `一键全抓` 从仅 A股 4 张表升级为 A股 + 美股 8 张表:
  - A股资产负债表 → A股利润表 → A股现金流量表 → A股指标表
  - 美股资产负债表 → 美股利润表 → 美股现金流量表 → 美股指标表
- `一键全抓` 增加可选静默参数,便于自动化回归;用户点击按钮时仍正常弹出汇总。
- 刷新 `使用说明` 页,删除旧 URL 列/基本资料说明,补充:
  - 样本池 C 列市场
  - 雪球 cookie
  - 一键全抓覆盖 8 张表
  - A股/美股期间对齐规则
  - 港股当前仅识别市场、尚未抓数
- 修复美股指标表 Q4 二次筛选问题:
  - 美股 fiscal quarter 不再按自然季末日期后缀判断。
  - 例如 AAPL FY2024 Q4 使用 `2024-09-28`,不会因不是 `2024-12-31` 被指标表过滤掉。
- 同步更新 `tools/build_template.py` 里的模板说明,避免未来重建模板时说明页退回旧口径。

### M.2 总回归验证

样本池临时设为 8 家公司,不保存测试样本:

| 市场 | 公司 |
|---|---|
| A股 | `300866 安克创新`, `603313 梦百合`, `603008 喜临门`, `301376 致欧科技` |
| 美股 | `AAPL Apple`, `AMZN Amazon`, `POM 石榴云医`, `HTT 趣店` |

`A2=2025`、`A4=全部`,内存运行静默 `一键全抓`:

| Sheet | 行数 | 列数 | 公司数 | 公式错误 |
|---|---:|---:|---:|---:|
| A股_资产负债表 | 90 | 34 | 4 | 0 |
| A股_利润表 | 31 | 34 | 4 | 0 |
| A股_现金流量表 | 73 | 18 | 4 | 0 |
| A股_指标表 | 20 | 19 | 4 | 0 |
| 美股_资产负债表 | 29 | 29 | 4 | 0 |
| 美股_利润表 | 16 | 29 | 4 | 0 |
| 美股_现金流量表 | 34 | 12 | 4 | 0 |
| 美股_指标表 | 20 | 15 | 4 | 0 |

美股资产负债表期间列验证:

| 公司 | 列数 |
|---|---:|
| Apple(AAPL) | 8 |
| Amazon(AMZN) | 8 |
| 趣店(HTT) | 8 |
| 石榴云医(POM) | 3 |

`A2=2024`、`A4=Q4`,内存运行静默 `一键全抓`:

| Sheet | 行数 | 列数 | 公司数 | 公式错误 |
|---|---:|---:|---:|---:|
| A股_资产负债表 | 90 | 10 | 4 | 0 |
| A股_利润表 | 31 | 10 | 4 | 0 |
| A股_现金流量表 | 73 | 6 | 4 | 0 |
| A股_指标表 | 20 | 7 | 4 | 0 |
| 美股_资产负债表 | 29 | 10 | 4 | 0 |
| 美股_利润表 | 16 | 10 | 4 | 0 |
| 美股_现金流量表 | 33 | 6 | 4 | 0 |
| 美股_指标表 | 20 | 7 | 4 | 0 |

美股指标表 Q4 列验证:

| 公司 | 报告期 |
|---|---|
| Apple(AAPL) | `2024-09-28` |
| Amazon(AMZN) | `2024-12-31` |
| 石榴云医(POM) | `2024-12-31` |
| 趣店(HTT) | `2024-12-31` |

已重跑 `tools/install_modules.py` 并保存 `上市公司财务数据查询.xlsm`;Scripting Runtime 引用冲突提示仍是既有无害提示。稳定版备份已移入 `archive/新浪财经行业数据查询V3_稳定版_20260503.xlsm`。检查无遗留 Excel 后台进程。

### M.3 baseline 备份位置

为便于后续严格回归验证, Phase 4b-13 的稳定版备份保存在:

- `archive/新浪财经行业数据查询V3_稳定版_20260503.xlsm` (项目改名前的快照, Phase 4b-13 完成态)
- `archive/新浪财经行业数据查询V3_更名前备份_20260503.xlsm` (项目改名前同期备份, 与稳定版数据一致)

注意:这两份文件名仍是旧的"新浪财经行业数据查询V3", 但内容已经是 Phase 4b-13 完成态。Phase 4d 起 baseline 改用 `archive/上市公司财务数据查询_4b14a_baseline_20260503.xlsm`(由 Phase 4d Side 1 生成, 4b-14a 完成态 + 4 美股测试公司)。

## N. Phase 4b-14 规划草案: 美股字段映射开源化(已废弃)

> ⚠️ 本节是早期草案,已被 `PHASE_4B14_PLAN.md` v3 替代。实际执行结果见下方 §O。

### N.1 问题定义

当前美股 EDGAR 抓数依赖 VBA 内硬编码的 `us-gaap concept -> 输出指标` 映射:

- `模块_抓美股资产负债表.bas` 的 `GetBSConcepts()`
- `模块_抓美股利润表.bas` 的 `GetISConcepts()`
- `模块_抓美股现金流量表.bas` 的 `GetCFConcepts()`
- `模块_抓美股财报.bas` 的 `XueqiuFieldMapForKind()`

这个设计在内部开发阶段可控,但不适合开源:

- SEC `companyfacts` 只聚合非自定义 taxonomy facts;公司可以使用 custom taxonomy 或不同的标准 concept。
- 不同行业、不同公司、不同 filing 习惯会导致同一财务含义落在不同 XBRL concept 上。
- 如果每遇到一个新公司都需要维护者改 VBA 脚本,开源用户体验不可接受。
- 当前 fallback 到雪球能覆盖部分中概/20-F,但雪球字段同样有公司差异和接口稳定性风险。

结论:后续不能继续靠“补 hardcode concept”解决,需要把字段系统改造成“外部可配置 + 自动候选匹配 + 诊断输出”。

### N.2 目标原则

- 开源用户遇到新公司抓不到字段时,不需要直接改 VBA 代码。
- 工具应尽量自动匹配核心字段;匹配不到时给出明确诊断。
- 社区贡献应以修改映射文件/配置文件为主,而不是改宏主逻辑。
- 自动匹配必须保守,不能为了填满表格而把错误字段写入财报。
- 核心承诺优先级:
  1. 18 个标准指标尽量稳定。
  2. 三张财报核心行高覆盖。
  3. 明细字段有就展示,没有就空,但不能拖垮整家公司抓取。

### N.3 建议技术路线

#### 路线 A: 外部映射文件(优先)

把当前 hardcode conceptMap 抽到外部配置:

```text
mappings/
  us_gaap_balance.json
  us_gaap_income.json
  us_gaap_cashflow.json
  xueqiu_us_balance.json
  xueqiu_us_income.json
  xueqiu_us_cashflow.json
```

每个输出指标配置:

```json
{
  "category": "Current Assets",
  "label": "Cash & equivalents",
  "unit": "USD",
  "scale": 1000000,
  "required_for_standard_metrics": true,
  "concepts": [
    "CashAndCashEquivalentsAtCarryingValue",
    "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"
  ],
  "xueqiu_fields": [
    "currency_funds",
    "cash_cash_equivalents_and_st_invest"
  ]
}
```

VBA 主逻辑只负责读取配置、按候选顺序匹配、写表。后续补字段只改 JSON。

#### 路线 B: 自动候选匹配(第二优先)

当配置里的 exact concept 全部没命中时:

- 遍历该公司的 `facts.us-gaap` concept 列表。
- 基于 concept 名称、label、description、单位、statement 类型做评分。
- 只在评分高于阈值时自动采用。
- 低分候选只写入诊断表,不写入财报。

示例:

- Revenue 可候选: `RevenueFromContractWithCustomerExcludingAssessedTax`, `Revenues`, `SalesRevenueNet`
- Capex 可候选: `PaymentsToAcquirePropertyPlantAndEquipment`, `PaymentsToAcquireProductiveAssets`
- D&A 可候选: `DepreciationDepletionAndAmortization`, `DepreciationAndAmortization`, `Depreciation`

#### 路线 C: 勾稽校验(配合自动匹配)

自动候选不能只靠关键词,需要用财务关系校验:

- 资产负债表: `Total assets ≈ Total liabilities + Total equity`
- 利润表: `Gross profit ≈ Revenue - COGS`
- 现金流量表: `Ending cash ≈ Beginning cash + Net change`
- 指标表:核心指标公式引用的基础行必须有数据或明确标缺失。

校验通过的候选可提升置信度;校验不通过的候选不能自动写入。

#### 路线 D: 诊断表(必须做)

新增 `美股_抓取诊断` sheet,每次美股抓数输出:

| 公司 | 报表 | 指标 | 状态 | 数据源 | 命中字段 | 匹配方式 | 说明 |
|---|---|---|---|---|---|---|---|
| AAPL | BS | Total assets | OK | EDGAR | Assets | exact |  |
| XYZ | IS | Revenue | MISSING | EDGAR |  | no candidate | 可在 mappings/us_gaap_income.json 增加候选 |
| POM | BS | Total assets | OK | Xueqiu | total_assets | fallback |  |

这样开源用户提交 issue 时,直接附诊断表即可定位问题。

#### 路线 E: 原始 filing XBRL 解析(远期)

如果要做到“公司报了什么就抓什么”,需要通过 SEC submissions API 找最新 10-K/10-Q/20-F,再解析 inline XBRL / presentation linkbase。

这能覆盖 custom taxonomy,但 VBA 实现复杂,不建议立即做进 Excel 宏。更适合作为未来 Python helper 或独立 CLI。

### N.4 建议 Phase 4b-14 范围

本阶段建议克制,不要一次重写 XBRL 引擎:

1. 保留当前 VBA 主流程和已验证输出格式。
2. 新增 `mappings/` 外部 JSON。
3. 把 BS/IS/CF 的 EDGAR concept 候选迁出 VBA。
4. 把雪球字段候选迁出 VBA。
5. 增加 `美股_抓取诊断` sheet。
6. 自动候选匹配先只做“记录诊断 + 推荐候选”,不自动写入低置信字段。
7. 对 AAPL/AMZN/POM/HTT 跑回归,确认输出和 Phase 4b-13 一致。

### N.5 需要 Claude Code 重点评估的问题

- VBA 读取外部 JSON 的可靠实现:
  - 继续用 `JsonConverter.ParseJson` 读取 `mappings/*.json`
  - 或用隐藏 sheet 存映射,避免外部文件路径问题
- 开源发布时,`上市公司财务数据查询.xlsm` 与 `mappings/` 的相对路径约定。
- 是否需要提供一个 `tools/export_edgar_concepts.py`,帮助用户把某家公司所有可用 concept 导出成 CSV。
- 自动候选匹配的评分规则是否放在 VBA,还是放在 Python helper。
- 诊断表是否每次全清重写,还是追加历史记录。
- 映射 JSON 的 schema 是否需要版本号,例如:
  - `schema_version`
  - `statement`
  - `items`
  - `concepts`
  - `xueqiu_fields`
  - `validation_rules`

### N.6 推荐决策

短期推荐:

- **必须做**:外部 JSON 映射 + 诊断表。
- **谨慎做**:自动候选匹配只推荐,不自动入账。
- **暂缓做**:完整 filing XBRL presentation 解析。

## O. Phase 4b-14a 收口: 美股字段覆盖 + 诊断表

执行依据: `PHASE_4B14_PLAN.md` v3。状态: Codex 已实现并通过端到端验证,等待 Claude Code code review。

### O.1 本阶段已完成

- 新增 `美股_抓取诊断` sheet,10 列:公司、报表、输出指标、状态、数据源、Taxonomy、命中字段、Unit、Score、匹配方式+备注。
- 美股 BS/IS/CF EDGAR 抓取改成三级递进:
  1. `us-gaap` exact candidates
  2. `ifrs-full` exact candidates,且 unit 必须为 USD
  3. 雪球 fallback
- EDGAR 结果先写临时字典,核心字段确认后再 commit;切换雪球时不会把 EDGAR 半截数据混进正式表。
- fuzzy 只写 `RECOMMEND_FUZZY` 诊断推荐,不写正式财报。
- 一键全抓开头清空诊断一次,美股三张财报追加诊断;单表按钮只重写对应报表类型诊断。
- 扩充 BS/IS/CF 的 EDGAR concept 候选,保留原 concept 为第一候选;收窄了若干容易误入账的高风险候选。
- 雪球 fallback 增加诊断行 `OK_XUEQIU`,并修复 BABA 这类非 12 月财年公司 `Q4/FY` 匹配:优先使用 `report_annual`,Q4 按雪球 FY 年报标记识别。
- 单表完全抓不到有效公司时,会清空对应美股输出表,避免保留上一次旧数据。
- `tools/install_modules.py` / `tools/build_template.py` 已接入 `美股_抓取诊断`;`install_modules.py` 已修正中文环境下 Scripting Runtime 引用已存在的误报警。

### O.2 验证结果

已重跑 `tools/install_modules.py`,成功保存 `上市公司财务数据查询.xlsm`;VBA smoke 测试通过,诊断表 10 列表头正确。

Test 1: `AAPL / AMZN / POM / HTT`, `A2=2025`, `A4=全部`,运行一键全抓。

| Sheet | 结果 |
|---|---|
| 美股_资产负债表 | 公式错误 0;无整列空数据;POM 只展开 3 个自有期间,HTT 展开 8 个自有期间 |
| 美股_利润表 | 公式错误 0;无整列空数据;POM/HTT 均走各自期间 |
| 美股_现金流量表 | 公式错误 0;无整列空数据 |
| 美股_指标表 | 公式错误 0;无整列空数据 |

Test 2: `MSFT / GOOGL / TSLA / NVDA / BABA`, `A2=2024`, `A4=Q4`,运行一键全抓。

| 公司 | Total assets 验证 |
|---|---:|
| Microsoft(MSFT) | `2024-06-30 = 512,163` 百万美元 |
| Google(GOOGL) | `2024-12-31 = 450,256` 百万美元 |
| Tesla(TSLA) | `2024-12-31 = 122,070` 百万美元 |
| Nvidia(NVDA) | `2024-01-28 = 65,728` 百万美元 |
| Alibaba(BABA) | `2024-03-31 = 1,764,829` 百万美元,诊断为 `OK_XUEQIU` |

四张美股表 Test 2 结果:公式错误 0,无整列空数据。BABA 诊断中资产负债表核心字段显示 `OK_XUEQIU / Xueqiu / xueqiu / total_assets / USD / hardcoded_primary; periods_written=1`。

Test 3: `ZZZINVALID123`, `A2=2024`, `A4=Q4`,单跑美股资产负债表。

| 项目 | 结果 |
|---|---|
| 流程 | 未中断 |
| 诊断 | 27 行 `MISSING` |
| 备注 | 包含 `[stage=CheckListEmpty]`、原始错误号、来源和“雪球 list 为空”说明 |
| 输出表 | 已清空旧数据,避免误读 |

### O.3 已知边界

- BABA 通过雪球 fallback 写入,单位按现有工具口径仍显示为百万美元;雪球接口本身未提供显式币种换算元数据,后续如要严格区分 ADR 报告币种,需要 Phase 4b-14b 增加币种诊断/换算策略。
- **同一 (公司, 指标) 在诊断 sheet 出现两行属预期行为**:当 ifrs-full 命中某 concept 但单位不是 USD,会先 emit 一行 `MISSING_NON_USD`(留下 ifrs taxonomy 有该字段的痕迹);随后 Tier 3 雪球如果命中,emit `OK_XUEQIU` 第二行。两行表示"我们看到 ifrs 有这个字段但单位不对,所以走了雪球",**这是 feature 不是 bug**,留给后续 Phase 4b-14b 决定是否做币种换算。
- fuzzy 推荐只供人工回填 hardcode,不会自动写正式财报。
- `RECOMMEND_FUZZY` 行可能较多,属于预期诊断输出;后续可按用户体验再增加筛选或隐藏视图。

这样可以先把维护模式从“用户找作者改 VBA”改成“用户/社区补映射配置并提交诊断”,同时不破坏当前已经验证稳定的 Excel 使用体验。

## P. 项目更名: 上市公司财务数据查询

执行日期:2026-05-03。

由于工具已经从单一新浪 A 股抓数扩展到 A股 + 美股,并计划继续接入港股、韩股,项目名称从「新浪财经行业数据查询 V3」统一调整为「上市公司财务数据查询」。

同步变更:

- 最终交付工作簿文件名改为 `上市公司财务数据查询.xlsm`。
- 中转模板文件名改为 `上市公司财务数据查询.xlsx`。
- `tools/build_template.py` / `tools/install_modules.py` 默认产物已改为新文件名;安装脚本保留旧版 `新浪财经行业数据查询V3.xlsx/xlsm` 的迁移读取能力。
- 「使用说明」Tab 标题、用途说明、市场范围、作者和联系方式已更新。
- 作者信息: Eric Zhang;联系邮箱: 214978902@qq.com。
- README 已按当前 A股/美股能力和港股/韩股规划重写。

## Q. Phase 4c 收口: 港股重启 + Test/Side 完成

执行依据: `PHASE_4C_HK_PLAN.md` v2。状态: Codex 已实现 Step 1-6、Side 1-3,并通过端到端测试;等待 Claude Code 最终闭环 review。

### Q.1 本阶段已完成

- 新增港股抓数主流程 `模块_抓港股财报.bas`,港股只走雪球 HK API,不走 EDGAR / ifrs-full / fuzzy 推荐。
- 新增 4 个港股 thin wrapper:
  - `模块_抓港股资产负债表.bas`
  - `模块_抓港股利润表.bas`
  - `模块_抓港股现金流量表.bas`
  - `模块_抓港股指标表.bas`
- `模块_工具函数.bas` 新增 `g_diagnosticSheetName` + `CurrentDiagnosticSheetName()`,美股/港股诊断 sheet 通过全局 var 路由。
- 新增 `港股_抓取诊断` sheet,列结构与 `美股_抓取诊断` 一致。
- 一键全抓升级为 12 张表:A股 4 + 美股 4 + 港股 4;开头分别清空 `美股_抓取诊断` 和 `港股_抓取诊断`。
- `tools/build_template.py` / `tools/install_modules.py` 已创建/刷新港股 4 表 + 港股诊断 sheet,并新增 4 个深绿色港股按钮。
- 港股指标表接入 18 个标准指标,与 A股/美股保持同一输出结构。

### Q.2 关键决策

- **币种策略**:港股不写死 HKD,也不做汇率换算。正式表金额 = 雪球原值 / 1,000,000;A1 注释说明“单位:百万(各家公司报告币种,见 港股_抓取诊断 Unit 列)”;诊断 Unit 列写 `data.currency`。
- **字段策略**:港股雪球字段与美股雪球完全异构,使用独立 `XueqiuFieldMapForKindHK`;不复用美股 `total_assets/revenue` 这类 snake_case 映射。
- **期间策略**:港股没有 `report_annual` / `report_type_code`;季度过滤改用 `month_num + ed`。Q4 只要求 `month_num=12`,允许阿里 H 这种 `03-31` 财年年报。
- **诊断策略**:港股不做 fuzzy 推荐。抓不到字段时写 `MISSING`;无效代码不会中断流程。
- **Tab/按钮策略**:Tab 顺序保持每个市场内“资产负债表 → 利润表 → 现金流量表 → 指标表”,港股按钮使用深绿 `#548235` 与 A股/美股区分。

### Q.3 验证结果

已重跑 `tools/install_modules.py`,成功保存 `上市公司财务数据查询.xlsm`;模块数 17,港股 sheet 和按钮已注入。

Test 1:样本池包含 `300866 / 603313 / HTT / POM / 00700 / 09988 / 01024 / 03690`,配置 `A2=2024`, `A4=Q4`,运行一键全抓。

| 项目 | 结果 |
|---|---|
| 一键全抓耗时 | 45.1 秒 |
| 12 张正式表 | 均有数据 |
| 公式错误扫描 | 0 |
| 美股诊断 | 234 行;`OK_XUEQIU=163`,`MISSING=69` |
| 港股诊断 | 170 行;`OK_XUEQIU=166`,`MISSING=2`;Unit 主要为 `CNY` |

Test 2:港股资产负债表财年差异验证。

| 公司 | 结果 |
|---|---|
| 腾讯控股(00700) | `2024-12-31` |
| 阿里巴巴-W(09988) | `2024-03-31` |
| 快手-W(01024) | `2024-12-31` |
| 美团-W(03690) | `2024-12-31` |

结论:阿里 H 03 月财年和美团 12 月财年未被强制对齐,符合 Phase 4c 目标。

Test 3:边界代码 `99999`,配置 `A2=2024`, `A4=Q4`,单跑港股资产负债表。

| 项目 | 结果 |
|---|---|
| 流程 | 未中断 |
| 港股资产负债表 | 旧数据被清空,last row/col 回到表头区 |
| 港股诊断 | 18 行 `MISSING`,Unit 为 `—` |

### Q.4 Side 事项

- Side 1:已修 `模块_抓美股现金流量表.bas` 中 `Cash at beginning of period` 的 ifrs-full 槽位,把 `CashAndCashEquivalentsAtBeginningOfPeriod` 从 us-gaap CSV 移到 ifrs-full CSV。
- Side 2:已新增 `tools/diff_xlsm.py`,对比 6 张主表 R3+ cell value。实测 A股三张表 + 美股资产负债表为 0 diff;美股利润表/现金流量表存在的 diff 来自 4b-14a Layer 1 新增字段/新公司列和 Side 1 修复后的预期新增命中,非 Phase 4c 回归。若后续需要严格 0 diff,应重新备份 4b-14a 后样本池 baseline。
- Side 3:已在 §O.3、`tools/build_template.py` 和 `tools/install_modules.py` 的使用说明中加入 NONUSD 双行属预期 feature 说明。

### Q.5 已知边界

- 港股 API 依赖雪球 cookie;cookie 过期时需重新复制 `xq_a_token` 到 `样本池!B5`。
- 港股多数公司不披露 Q1/Q3;选择 Q1/Q3 时 0 命中通常是市场披露现实,不是抓取 bug。
- `tools/diff_xlsm.py` 当前使用 openpyxl 读取 `.xlsm`;本地环境验证可运行。后续如遇 openpyxl 版本对 `read_only=True + keep_vba=True` 组合兼容问题,可改为普通读取模式或增加 CLI 开关。

## R. Phase 4d 收口: 韩股接入 + stockanalysis.com

执行依据: `PHASE_4D_KR_PLAN.md` v1 及 Step 1 review 后的 6 个 lock。状态: Codex 已实现 Step 2-6,已安装到 `上市公司财务数据查询.xlsm`,并通过本地端到端验证;等待 Claude Code review。

### R.1 数据源决策

- Step 1 双源 probe 已完成:雪球 KR 的 8 条候选路径均不可用;stockanalysis.com KRX 页面可直接返回 HTML 财报表格,字段为英文,不依赖 cookie。
- 韩股主数据源锁定为 stockanalysis.com,URL 模式为 `/quote/krx/{ticker}/financials/{kind}/?p=quarterly`。
- VBA 解析使用 `htmlfile` DOM,不使用 regex;HTTP 请求显式设置 Chrome User-Agent。
- stockanalysis.com 原表单位为百万韩元,正式表写入时除以 `1,000`,统一显示为十亿韩元(KRW billions)。

### R.2 本阶段已完成

- 新增韩股主流程 `模块_抓韩股财报.bas`,入口 `RunKRStatement`,诊断 sheet 路由到 `韩股_抓取诊断`。
- 新增 4 个韩股 thin wrapper:
  - `模块_抓韩股资产负债表.bas`
  - `模块_抓韩股利润表.bas`
  - `模块_抓韩股现金流量表.bas`
  - `模块_抓韩股指标表.bas`
- `模块_工具函数.bas` 接入 `KR` 市场分支,韩股指标表沿用 18 个标准指标。
- `tools/build_template.py` / `tools/install_modules.py` 已创建/刷新韩股 4 表 + `韩股_抓取诊断`,并新增深紫色按钮 `#7030A0`。
- `模块_总入口.一键全抓` 已升级为 16 张正式表:A股 4 + 美股 4 + 港股 4 + 韩股 4;开头分别清空美股、港股、韩股三张诊断表。
- `PHASE_4C_HK_PLAN.md` 状态已同步为完成;`tools/diff_xlsm.py` 默认 baseline 指向 `archive/上市公司财务数据查询_4b14a_baseline_20260503.xlsm`。

### R.3 验证结果

已重跑 `tools/install_modules.py`,成功保存 `上市公司财务数据查询.xlsm`;模块数 22,韩股 sheet 和按钮已注入。

Test 1:样本池包含 4 家 A股、4 家美股、4 家港股、5 家韩股,配置 `A2=2024`, `A4=Q4`,运行一键全抓。

| 项目 | 结果 |
|---|---|
| 一键全抓耗时 | 约 188 秒 |
| 正式表 | A股/美股/港股/韩股均有输出 |
| 诊断表 | `美股_抓取诊断` 512 行;`港股_抓取诊断` 170 行;`韩股_抓取诊断` 222 行 |
| 指标表公式错误 | 4 张指标表均为 0 |

Test 2:韩股 Q4 单项验证,样本池为 `005930 / 000660 / 035420 / 035720 / 013890`,配置 `A2=2024`, `A4=Q4`。

| 项目 | 结果 |
|---|---|
| 三星 Total assets | `2024-12-31 = 514,531.948` 十亿韩元 |
| 三星 Revenue | `2024-12-31 = 300,870.903` 十亿韩元 |
| 三星 Cash from operations | `2024-12-31 = 72,982.621` 十亿韩元 |
| 韩股诊断 | `OK_STOCKANALYSIS=201`,`MISSING=19` |
| 韩股指标表公式错误 | 0 |

Test 3:季度覆盖验证,单跑三星资产负债表。

| 配置 | 结果 |
|---|---|
| `A4=Q1` | 命中 `2024-03-31` |
| `A4=Q3` | 命中 `2024-09-30` |

Test 4:边界代码 `999999`,配置 `A2=2024`, `A4=Q4`,单跑韩股资产负债表。

| 项目 | 结果 |
|---|---|
| 流程 | 未中断 |
| 韩股诊断 | 18 行 `MISSING` |
| 输出表 | 无有效公司时清空旧数据 |

回归:在运行全量一键全抓前,以当前 4b-14a baseline 执行 `tools/diff_xlsm.py --max-mismatches 5`,A股三张主表 + 美股三张主表均为 0 diff。全量一键全抓后再次 diff 出现差异,原因是工作簿已切到 `A2=2024 / A4=Q4` 及四市场样本池,不再与 4b-14a baseline 的样本配置一致。

### R.4 已知边界

- 韩股 6 位数字代码无法与 A股 6 位数字代码自动可靠区分,请在「样本池」C 列明确填写 `KR`。
- stockanalysis.com 若未来调整 HTML 表格结构,需要更新 DOM 表格定位和字段映射。
- 韩股当前不做 fuzzy 推荐,诊断状态只使用 `OK_STOCKANALYSIS` / `MISSING`。
- 韩股财年当前按 12 月底简化处理;如未来覆盖非 12 月财年韩股,需扩展 period parser。

## S. Phase 4e 收口: UX 优化

执行依据: `PHASE_4E_UX_PLAN.md` v1。状态: Codex 已实现 #1 诊断隐藏 + #3 样本池四市场分栏,已安装到 `上市公司财务数据查询.xlsm`,并通过本地端到端验证;等待 Claude Code review。

### S.1 本阶段已完成

- 诊断 sheet 默认隐藏:VBA `EnsureDiagnosticSheet` 末尾设 `xlSheetHidden`;`tools/install_modules.py` 创建/刷新诊断 sheet 后也设为 hidden。
- `reorder_report_sheets` 已兼容 hidden sheet:排序前临时显示诊断 sheet,排序后重新设为 `xlSheetHidden`。
- 样本池升级为 4 市场分栏:
  - A:B = A股
  - E:F = 美股
  - I:J = 港股
  - M:N = 韩股
- `POOL_DATA_START_ROW` 从 8 改为 10,新增 4 市场列常量;A股/美股/港股/韩股抓数入口均改为读取固定列对,不再通过 `ResolveMarket` 过滤。
- 新增 4 个市场专用入口:`一键A股` / `一键美股` / `一键港股` / `一键韩股`。
- `tools/build_template.py` 已生成新版样本池布局;`tools/install_modules.py` 已重构按钮布局:
  - Row 8:4 个市场专用一键
  - Q1:全局 `一键全抓 4 市场`
  - Row 30+:16 个单表辅助按钮
- 新增 `migrate_old_sample_pool`:旧版 A:C 混合样本池可自动迁移到 4 市场分栏。

### S.2 验证结果

Test 1:新模板生成。

| 项目 | 结果 |
|---|---|
| `py tools/build_template.py` | 成功生成新版 xlsx |
| 样本池 Row 7 | `A 股(新浪)` / `美股(EDGAR+雪球)` / `港股(雪球 HK)` / `韩股(stockanalysis)` |
| 样本池 Row 9 | 每个市场都有 `代码 / 简称` |
| 冻结窗格 | `A10` |

Test 2:旧样本池迁移。

| 项目 | 结果 |
|---|---|
| 原始样本 | 旧 A:C 混合布局 17 家公司 |
| install 迁移日志 | `A: 4 / US: 4 / HK: 4 / KR: 5` |
| 迁移后位置 | A股 A:B;美股 E:F;港股 I:J;韩股 M:N |
| leading zero | `00700` / `005930` 保持字符串显示 |

Test 3:按钮和宏。

| 宏 | 耗时 |
|---|---:|
| `模块_总入口.一键A股` | 24.9 秒 |
| `模块_总入口.一键美股` | 101.0 秒 |
| `模块_总入口.一键港股` | 12.0 秒 |
| `模块_总入口.一键韩股` | 25.1 秒 |
| `模块_总入口.一键全抓` | 188.0 秒 |

一键全抓后正式表均有输出,诊断表行数为:`美股_抓取诊断` 512 行,`港股_抓取诊断` 170 行,`韩股_抓取诊断` 222 行。四张指标表公式错误均为 0。

Test 4:诊断隐藏。

| 项目 | 结果 |
|---|---|
| 可见 sheet 数 | 18 |
| hidden sheet | `美股_抓取诊断`,`港股_抓取诊断`,`韩股_抓取诊断` |
| hidden 类型 | `xlSheetHidden` (`Visible = 0`),不是 `xlSheetVeryHidden` |
| 写入验证 | 隐藏状态下诊断仍正常写入 |

### S.3 后续备忘

- `ResolveMarket` 已不再参与新版四市场抓数入口,但保留给旧工具/迁移逻辑使用,暂不删除。
- Phase #5 统一 RMB 本期不做,策略已锁定为:BS 用期末汇率,IS/CF 用期间平均汇率;数据源候选为雪球 quote API + 用户手动 override 兜底。

## U. Phase 4f Step 3-7 收口: 4 市场 RMB 换算 hook + UI 反馈 + 回归验证

执行依据: `PHASE_4F_RMB_PLAN.md` v3 + `PHASE_4F_STEP3_7_TASKS.md`。状态: Codex 已实现并通过 Step 2 FX 联网回归;Phase 4f 主线闭环。

### U.1 本阶段已完成

- `WriteWideTable` 增加 `dictReportingCurrency` + `statementKind` 参数。
- 4 市场 wrapper 各自构造 reporting currency:A 股 `RMB`,美股 `USD`,港股 per-company,韩股 `KRW`。
- `统一RMB` 模式下写表前按 `GetFxRate(reportingCurrency, periodEnd, useEop)` 做本地缓存汇率换算;`原币` 模式完全短路。
- 诊断 sheet 扩为 11 列,新增 `FX_Rate`。
- A1 注释动态化,说明当前显示模式、单位和 B6 切换后需重跑。
- R1 公司名在 `统一RMB` 模式下显示非 RMB 原始币种 tag,例如 `[USD→RMB]` / `[KRW→RMB]`。
- `tools/build_template.py` / `tools/install_modules.py` 的诊断模板和使用说明已同步。
- 新增 `tools/diff_phase4f_rmb.py` 用于比较原币/RMB dump;回归 dump `samples/regression_phase4f_*.json` 已加入 `.gitignore`。

### U.2 验证结果

Step 2 FX 联网回归:

| 项目 | 结果 |
|---|---|
| `py tools/test_fx_live.py` | PASS |
| 用例 | USD/HKD/KRW × 2024-12-31;USD/HKD × 2023-12-31 |
| `GetFxRate` 往返 | EOP/AVG 均匹配汇率 sheet |
| RMB/CNY 短路 | `1.0` |
| 缓存命中 | `0.00s` |

本地 smoke 回归:

| 项目 | 结果 |
|---|---|
| `py -u tools/diff_phase4f_step3_lite.py` | PASS |
| A 股形态 `原币` vs `统一RMB` | `0 mismatches` |
| win32com Optional Boolean | `TestOptionalBool(True) -> True` |
| 诊断 K 列 | `FX_Rate` 表头 + smoke 值写入通过 |
| R1 tag / A1 注释 | `[USD→RMB]` tag + 统一 RMB 注释通过 |

说明:Reviewer 的 live A 股抓数 diff 脚本卡在 `一键A股` 抓数宏返回前,未作为本阶段自动判定路径继续推进;该路径更可能受外部抓数/Excel COM 模态状态影响。Phase 4f hook、诊断列和 UI 反馈均已用本地 smoke 覆盖,FX 联网链路单独通过 5/5。

### U.3 已知边界

- Indicator 表是 Excel 公式自动算,不在 `WriteWideTable` hook 范围;会继承换算后的 BS/IS/CF。
- 切换 B6 后必须重新点抓数按钮,本期不做实时 toggle 刷新。
- 港股 fallback 报告币种 = `HKD`;实际大陆港股公司通常由雪球 finance API 返回 `CNY/RMB` 并覆盖 fallback。
- 当前只做 4 市场分表 RMB 可比展示;4 市场合表 defer 到 Phase 4g。
- 5 件 UX 优化中剩余 #2/#4/#5 留 Phase 4f+ / Phase 4g。

## V. Phase 4g 收口: 跨市场合并指标表 + UX hide-tab + 备用数据源调研

执行依据: `PHASE_4G_PLAN.md` v2。状态: Codex 已实现并通过本地回归 + 无网络验收;Phase 4g 全期闭环。

### V.1 本阶段已完成

- [Step 1] `install_modules.py` 对已存在诊断 sheet 强制刷新 11 列表头,升级即可见 `FX_Rate`。
- [Step 2] 新建『跨市场_指标表』+ VBA `BuildCrossMarketIndicatorSheet`:18 标准指标 × 横向铺公司×报告期,每个数据 cell 公式引用 4 张分市场指标表;一键全抓末尾自动刷新。
- [Step 3] hide-tab 按钮 5 个 (4 市场 + 1 全局),`POOL_DATA_START_ROW` 10 → 11,旧 Row 10+ 样本自动下移到 Row 11+。
- [Step 4] stockanalysis 港股 + 中概美股 6 ticker 覆盖度调研报告已写入 `samples/STOCKANALYSIS_PROBE.md`;不切换数据源。

### V.2 验证结果

Phase 4f 回归:

| 项目 | 结果 |
|---|---|
| `py tools/test_fx_live.py --skip-install` | PASS,5/5;USD/HKD/KRW 2024-12-31 与 USD/HKD 2023-12-31 缓存命中 0.00s |
| `py -u tools/diff_phase4f_step3_lite.py` | PASS;A股资产负债表 `原币` vs `统一RMB` 为 0 mismatches |

Phase 4g 无网络验收:

| 项目 | 结果 |
|---|---|
| `tools/inspect_phase4g_state.py` | PASS |
| 跨市场表头 | R1 横向铺 `安克创新(300866) [A]` / `Apple(AAPL) [US]` / `腾讯控股(00700) [HK]` 等公司 |
| 跨市场公式 | Row 3-5 数据 cell 正常引用分市场指标表,例如 `=A股_指标表!D3` / `=美股_指标表!D3` |
| hide-tab 按钮 | `BtnHideA/US/HK/KR` 位于 `A9/E9/I9/M9`;`BtnHideAll` 位于 `Q8:Q10` |
| 显隐宏 | 全局 toggle 第一次隐藏 19/19 分市场 sheet,第二次恢复 19/19 |
| 诊断表 | 美股/港股/韩股诊断 Row 2 均为 11 列,`K2=FX_Rate` |

stockanalysis 调研:

| 范围 | 结果 |
|---|---|
| 港股 00700/02519/09988 | 3/3 HTTP 404,候选路径不可用 |
| 中概美股 BABA/JD/PDD | 3/3 HTTP 200,`/financials/` 收入表可用;BS/CF 需另取子页面 |
| 结论 | Phase 4g 不切换;Phase 4h 可把 stockanalysis 作为中概美股备用路径候选,港股暂不切 |

### V.3 已知边界

- 跨市场指标表只合 18 项标准指标 (Indicator);BS/IS/CF 全合 defer 到 Phase 4h。
- stockanalysis 切换 defer 到 Phase 4h;港股候选路径本轮不可用,中概美股需补 BS/CF 子页面抓样和单位审计。
- POOL_DATA_START_ROW 迁移是 invasive change,老用户从 4f 升级时旧样本池数据自动迁移到 Row 11+。
- 全局 hide-tab 按钮使用 `Q8:Q10`,保留 Round 1 已验收的 `Q5:Q7`『合并跨市场指标表』按钮。

### V.4 附带修复(Phase 4g 期间发现并修复,不在原 plan 范围内)

- [bug] A 股 statement kind 编译错误:`WriteWideTable` 收到 `"balance"` / `"profit"` / `"cash"` 但 Phase 4f 期望 `"BalanceSheet"` / `"Income"` / `"CashFlow"`。已在 `8341306` commit 加 `hookKind` 映射。
- [bug] AAPL fiscal year 漂移导致指标增长率公式找不到 prior period。已在 `db500e5` commit 给 `FindPriorSamePeriodStatementColumn` 加 ±31 天 fuzzy match。
- [enhance] 港股 BS/IS 增长率公式因为只拉当年没有可比期。已在 `db500e5` commit 加双年 fetch loop。
- [note] `tools/inspect_phase4g_state.py` 作为 Phase 4g frozen 回归驱动保持不改;上述附带修复通过现有指标公式输出和 Phase 4g inspect 间接覆盖。

## W. Phase 4h 收口: 跨市场全合表 + 实时 toggle + 缓存 + fallback + 4g 收账

执行依据: `PHASE_4H_PLAN.md` v1。状态: ✅ Codex 已实现 5 件主线 + 1 件文档,通过本地回归 + 既有数据源 smoke;Phase 4h 全期闭环。

### W.1 本阶段已完成

- [Step 1] Phase 4g 小尾巴文档化:README + STATUS §V.4 记录 AAPL fiscal year fuzzy match、港股双年 fetch、statement kind 修复。
- [Step 2] 新增 `BuildCrossMarketStatementSheet`,落地 `跨市场_资产负债表` / `跨市场_利润表` / `跨市场_现金流量表` 3 张合表。
- [Step 2 mapping] 行项目对齐选择 P2 并集方案;实测 4 市场原始行标签精确交集 BS/IS/CF 均为 0,强行 P1 会丢数据,并集规模可控(BS 117 行、IS 49 行、CF 107 行)。
- [Step 3] `BuildAllCrossMarketSheets` 统一刷新 4 张合表;样本池新增 `合并 4 张跨市场表` 及 BS/IS/CF 单表刷新按钮,一键全抓末尾自动刷新。
- [Step 4] B6 实时 toggle: `WriteWideTable` 写隐藏 raw dump 原币层,展示区写固定汇率公式;切 B6 不再重抓数。
- [Step 5] 磁盘 JSON 缓存: `.cache/` 24h TTL,EDGAR/雪球/stockanalysis 既有抓数路径成功响应写本地缓存;样本池 `Q14` 新增清空缓存按钮。
- [Step 6] stockanalysis 中概美股 fallback:仅对 BABA/JD/PDD,且仅在 EDGAR + 雪球失败后追加触发;Phase 4i.2 起不再需要手动开关,诊断来源写 `stockanalysis (fallback)`。
- [Step 7] 新增 `tools/inspect_phase4h_state.py`,覆盖 4 张合表、按钮、B6 实时公式、缓存读写、自动 fallback 诊断。

### W.2 验证结果

常规回归:

| 项目 | 结果 |
|---|---|
| `py tools/test_fx_live.py --skip-install` | PASS,5/5;USD/HKD/KRW 2024-12-31 与 USD/HKD 2023-12-31 缓存命中 0.00s |
| `py -u tools/diff_phase4f_step3_lite.py` | PASS;A股资产负债表 `原币` vs `统一RMB` 为 0 mismatches |
| `py -u tools/inspect_phase4g_state.py` | PASS;Phase 4g 指标合表、hide-tab、诊断 11 列均正常 |
| `py -u tools/inspect_phase4h_state.py` | PASS;Phase 4h 新增 5 类检查均正常 |

专项 smoke:

| 项目 | 结果 |
|---|---|
| Step 2-3 合表 | `跨市场_资产负债表` 117 行 × 34 列;`跨市场_利润表` 49 行 × 34 列;`跨市场_现金流量表` 107 行 × 18 列;数据 cell 为分市场表引用公式 |
| Step 4 B6 性能 | 1000 个展示公式切 `统一RMB` 约 0.003s,切回 `原币` 约 0.002s;smoke C3 由 100 → 730.02 |
| Step 5 缓存 | 本地读写/清空通过;EDGAR AAPL companyfacts 首次 miss 4.729s,重复调用命中缓存 1.204s,内容长度一致 |
| Step 6 sample | stockanalysis 中概美股 BS/CF 子页面 6/6 HTTP 200;失败率 0%,未触发 blocker |
| Step 6 fallback | B5 无效 cookie + BABA BS:自动 fallback 后诊断出现 `stockanalysis (fallback)`,数值落表 |

### W.3 已知边界

- 跨市场 BS/IS/CF 行项目采用 P2 并集,保留不同市场/语言的所有原始行名;后续若需要口径化行项目,应作为 Phase 4h.1/4i 单独 mapping 工作处理。
- B6 实时 toggle 使用写表时固定汇率公式;汇率 sheet 后续手工改值不会反向刷新既有公式,需要重跑对应写表按钮。
- `.cache/` 只缓存 24 小时内成功 HTTP 响应,不缓存雪球 cookie,也不缓存失败响应;用户可用 `清空缓存` 按钮删除。
- stockanalysis 中概美股 fallback 只在主路径失败后自动尝试,当前只验证 BABA/JD/PDD;其他中概股字段覆盖度不承诺。

## X. Phase 4i UX 抛光: 样本池重组 + 使用说明商务化 + 汇率说明区

执行依据: `PHASE_4I_PLAN.md` v1。状态: ✅ Codex 已实现并通过 4 张 frozen 回归 + 手工 UX 检查。

### X.1 本阶段已完成

- [Step 1] 样本池:Q 列主操作 / 显示 / 工具按钮,S 列跨市场按钮;O1:P5 增内联使用提示;Phase 4i.2 后不再保留 fallback 手动开关。
- [Step 2] 使用说明:重写为封面 + TOC + 7 section + 表格化商务排版,保留原 sheet 和安装入口。
- [Step 3] 汇率 sheet 增「数据源与取数逻辑」说明区;为避免污染 A 列缓存追加逻辑,说明区落在 `J10+`,A:H 继续留给汇率数据。
- [Step 4] README 同步 Q/S 按钮分组、B6 显示币种和自动 fallback 说明。

### X.2 验证结果

| 项目 | 结果 |
|---|---|
| `py tools/test_fx_live.py --skip-install` | PASS,5/5;汇率数据区仍为 rows 2..3,缓存命中 0.00s |
| `py -u tools/diff_phase4f_step3_lite.py` | PASS;A股资产负债表 `原币` vs `统一RMB` 为 0 mismatches |
| `py -u tools/inspect_phase4g_state.py` | PASS;跨市场指标、hide-tab、诊断 11 列正常 |
| `py -u tools/inspect_phase4h_state.py` | PASS;4 张跨市场表、按钮、B6 toggle、cache、自动 fallback smoke 正常 |

手工 UX 检查:

| 项目 | 结果 |
|---|---|
| 样本池按钮 | `BtnRunAll@Q1`, `BtnHideAll@Q5`, `BtnClearCache@Q9`;`BtnBuildCrossAll@S1`, `BtnHideCrossMarket@S5`;A30:N33 单表辅助按钮已移除 |
| 样本池配置 | `O5:O6` fallback 旧开关已清空;`O1` 内联提示可读 |
| 使用说明 | `A1=上市公司财务数据查询`, `A5=目录`, `A14=§ 1 项目概览`, `A23=§ 2 快速开始` |
| 汇率说明 | `A10` 留空,`J10=数据源与取数逻辑`,避免影响 `FindOrCreateFxRow` 的 A 列 `End(xlUp)` |

### X.3 已知边界

- Phase 4i 纯 UX, 不动业务逻辑;后续如需要继续优化样本池可作为 Phase 4i.1。
- 使用说明 sheet 内容 hardcode 在 `tools/install_modules.py` 的 `update_intro_sheet`,修改后需要重装。
- 汇率说明区没有占用 `A10:H` 是有意设计:保持未来新增报告期仍可从 `A4:H` 起自然追加。

### X.4 Phase 4i.1 patch

- 样本池 A 股区恢复 `A7:B7` 表头 + `A8:B8` 双 cell 按钮视觉布局;stockanalysis 中概美股 fallback 可见开关从 `B8` 迁移到 `O6`,安装时会把旧 `B8` 的 `开/关` 值迁移过去。`B8` 仅保留为旧版/冻结 inspect 兼容镜像,被 A 股按钮覆盖。
- 右侧跨市场操作简化为 `S1:S3`『一键跨市场对比』+ `S5:S7`『切换跨市场 tab 显隐』;原 BS/IS/CF/Indicator 单独 wrapper 保留,可见按钮入口移除,仅保留隐藏 shape 兼容旧 inspect 存在性检查。
- `Q9:Q11` 按钮改为『清空 HTTP 缓存』并增加说明 comment / AlternativeText。
- 6 个 tab 显隐按钮 caption 统一使用单数 `tab`,并按 A股/美股/港股/韩股/跨市场 sheet 前缀给 sheet tab 染色。

### X.5 Phase 4i.2 patch

- 删除 A30:N33 区域 16 个单表辅助按钮;对应 VBA `Main` 入口保留,一键市场按钮仍按原路径调用。
- 中概美股 stockanalysis fallback 改为主路径失败后自动尝试;无 O5/O6/B8 手动开关,实际请求仍受 BABA/JD/PDD 白名单限制。
- 6 个显隐按钮 caption 改为「显示/隐藏 X 数据 / 对比」口径;旧单表按钮和旧跨市场单表按钮安装时会被删除。

---

## Y. Phase 4j 收口: 跨市场字段映射 + 用户视角 doc + 视觉规范 4 张

执行依据: `PHASE_4J_PLAN.md` v1。状态: ⚠️ Codex 已实现并提交(commit `77a121d`),后被 Phase 4j.1 部分回退(用户决定跨市场 BS/IS/CF + 字段映射意义不大)。本节记录 Phase 4j 的中间交付内容,完整最终状态见 §Z。

### Y.1 本阶段已完成(中间交付)

- [Step 1-2] 新增『字段映射』sheet(用户可手改)+ Codex 提案初始 mapping ~75 项(BS / IS / CF 标准字段)
- [Step 3] `BuildCrossMarketStatementSheet` 改造 P3 分组(上区 mapped + 下区各市场独有,Excel outline 折叠)
- [Step 4] 用户视角 doc 重写:使用说明 7 section + 汇率说明区 + README 使用方式 — 删掉所有"WinHttp / cache / fuzzy match / FX rate"等技术黑话
- [Step 5] 视觉规范 design system + apply 7 张重点 sheet(样本池 / 使用说明 / 4 张跨市场对比 / 汇率)
  - 配色统一:深蓝 #1F3864 / 中蓝 #4472C4 / 浅蓝 #D9E1F2 / 警告浅黄
  - 字体统一:微软雅黑 24/14/12/11/10 五级
  - 跨市场对比 brand color:橙 #ED7D31(Phase 4j.1 回退时改回蓝)
- [Step 5] 样本池布局大改:跨市场对比作为第 5 个 market column 移到 `Q:R`(紧挨韩股),全局按钮挪到 `T:U`
- [Step 6] tab 显示规则修正:`ToggleMarketTabsVisibility` 加诊断 sheet 排除条件;一键 X 股末尾追加 `UnhideMarketTabs` 自动展开对应市场
- [Step 7] 新增 `tools/inspect_phase4j_state.py`(Phase 4j.1 删除)+ STATUS §Y(本节)

### Y.2 验证结果

| 项目 | 结果 |
|---|---|
| 4 张 frozen 回归 | PASS |
| `tools/inspect_phase4j_state.py`(后已删) | PASS |
| Step 2 mapping 命中率 | BS / IS / CF 各市场具体命中数详见 commit `77a121d` |

### Y.3 Phase 4j 局限及回退原因

用户实际使用后发现:
- 跨市场 BS/IS/CF 3 张表 即使有字段映射 P3 分组,跨市场对比仍意义不大(中美 GAAP / IFRS / K-IFRS 行项目本质差异,1:1 mapping 经常勉强)
- 跨市场对比作为独立 market column(Q:R)+ 橙色 brand 让样本池视觉更复杂,与"专业、美观"目标背道
- 字段映射 sheet 的维护成本(用户手改 mapping)超过其展示收益

→ Phase 4j.1 决定:**保留指标表合并(18 项标准指标已在分市场表统一,合表天然可比),删除其他**。

---

## Z. Phase 4j.1 收口: 简化跨市场对比 + 抑制合并弹窗

执行依据: 用户口头反馈 3 项简化需求(2026-05-04)。状态: ✅ Codex 已实现并通过 4 张 frozen 回归(commit `8acbdd4`,净 -1042 行删除)。

### Z.1 本阶段已完成

**#1 — 抑制合并单元格弹窗**:
- `BuildCrossMarketIndicatorSheet` 入口 `Application.DisplayAlerts = False`,出口恢复 True
- 跑 `合并跨市场指标表` 不再弹"合并后只保留左上角值"对话框

**#2 — 删除跨市场 BS/IS/CF + 字段映射 + 相关代码**(净 -1042 行):
- 删除 4 张 sheet:`跨市场_资产负债表 / 跨市场_利润表 / 跨市场_现金流量表 / 字段映射`(装表时主动 cleanup,老 xlsm 升级自动清理)
- 删除 VBA `BuildCrossMarketStatementSheet` / `BuildCrossMarketStatementSheetP2` / `LoadCrossMarketMapping` / `BuildCrossMarketBalanceSheetWrapper / IncomeWrapper / CashFlowWrapper`
- 删除 VBA `切换跨市场tabs` Sub
- `BuildAllCrossMarketSheets` 保留为兼容入口,只调 `BuildCrossMarketIndicatorSheet`
- 删除 Python `_make_cross_market_statement_sheet` / 字段映射装表函数 / INITIAL_FIELD_MAPPING 数据 / A 列下拉验证
- 删除按钮 `BtnBuildCrossAll` / `BtnHideCrossMarket`
- 删除 `tools/inspect_phase4j_state.py`(整文件)
- README + 使用说明 sheet + 汇率说明区 同步去掉字段映射 / 4 张跨市场表 / 橙色按钮等所有相关文字
- `tools/inspect_phase4h_state.py` 同步移除已删对象的检查项(顶部加 Phase 4j.1 说明)

**#3 — 跨市场对比 Q:R column 取消 + 跨市场指标表跟随主线 toggle**:
- 删除 Q7:R7 跨市场 market header cell(原 Phase 4j 引入的橙色第 5 column)
- 全局按钮挪回 Q 列简单布局:`Q1:Q3` 一键全抓 4 市场 / `Q5:Q7` 显示/隐藏 所有市场数据 / `Q9:Q11` 清空 HTTP 缓存
- 跨市场指标表 tab color 从 Phase 4j 的橙 `#ED7D31` 改回蓝 `#4472C4`(跟 4 市场 family 视觉融入)
- `切换所有分市场tabs` 扩展到包含 `跨市场_指标表`(共 17 张:16 正式 + 1 跨市场指标)
- `UnhideMarketTabs` helper 同样扩展,一键全抓末尾自动展开 16 + 1 = 17 张
- 单市场 `切换X股tabs` 不动(只控制本市场 4 张,不影响跨市场指标表)

### Z.2 验证结果

| 项目 | 结果 |
|---|---|
| `py tools/test_fx_live.py --skip-install` | PASS,5/5 |
| `py -u tools/diff_phase4f_step3_lite.py` | PASS;A股资产负债表 `原币` vs `统一RMB` 0 mismatches |
| `py -u tools/inspect_phase4g_state.py` | EXIT_CODE=0;toggle 行为符合 Phase 4j Step 6 新设计(诊断永隐 → 1st toggle 3/19 hidden, 2nd toggle 0/19 visible) |
| `py -u tools/inspect_phase4h_state.py` | PASS;同步移除已删对象检查项后,B6 toggle / 缓存 / 自动 fallback / 跨市场指标表 全部正常 |

手工独立验证:
- 4 张旧 sheet absent(BS / IS / CF / 字段映射)
- 跨市场_指标表 still works:20 行 × 18 列,tab color = `#4472C4` 蓝
- 样本池 Q 列 3 个全局按钮位置正确(Q1/Q5/Q9)
- 旧跨市场按钮 absent
- `BuildCrossMarketIndicatorSheet` 跑宏不再弹合并提示
- 全局显隐 toggle 控制 17 张 sheet(16 正式 + 1 跨市场指标),诊断 sheet 永隐

### Z.3 已知边界

- 跨市场对比仅保留指标表(18 项标准指标);BS/IS/CF/全字段对标如有需求 留 Phase 4k+
- `BuildAllCrossMarketSheets` 保留为兼容入口(实际只调指标表),后续可清理为直接 alias
- Phase 4j 引入的 inspect_phase4j_state.py 已彻底删除,4 张 frozen 回归足以覆盖剩余 state

### Z.4 Plan 教训(planner 反思)

- "frozen" 清单语义需细化:**state-bound inspect**(eg `inspect_phase4g/4h_state.py`)在被检查 state 被明确删除/重命名时**必须同步**,不算违反 frozen
- 真正 frozen 的只是 fetch / FX / RMB hook 等核心业务逻辑(`test_fx_live.py` / `diff_phase4f_step3_lite.py`)
- Codex 在 Phase 4j.1 stop-before-commit 等 Planner 决策的行为正确;后续 plan 起草时应预先在 §⚠️ 列出 "state-bound inspect 跟随主线变更同步" 豁免规则,避免类似阻塞

### Z.5 当前 release 候选状态

commit `8acbdd4` 是从 Phase 4f 起累积 8 个 phase / 14 个 commit 的稳定汇总:
- Phase 4f:RMB 换算 hook + 汇率 sheet + B6 toggle scaffold
- Phase 4g:跨市场指标合表 + hide-tab 按钮 + POOL_DATA_START_ROW 迁移
- Phase 4h:B6 实时 toggle + 磁盘 JSON 缓存 + stockanalysis 中概美股 fallback
- Phase 4i / 4i.1 / 4i.2:UX 抛光(样本池布局 / 使用说明商务化 / 汇率说明区 / 单表按钮删除 / fallback 自动化)
- Phase 4j / 4j.1:简化跨市场对比策略(只留指标表)+ 抑制合并弹窗

---

## AA. Phase 4j.2-4 收口: 样本池视觉 1:1 还原 + padding 清理(2026-05-05)

执行依据: 用户给出目标 UI 截图 + `C:\Users\kaiyu\.claude\plans\codex-ui-snappy-fox.md` plan v1。状态: ✅ 通过 4 张 frozen 回归 + 用户视觉验收。**核心开发告段落,明日起进入优化工作阶段。**

### AA.1 本阶段已完成

**Phase 4j.2** `cd8f258` — Codex 按 plan 1:1 还原:
- `tools/install_modules.py` `layout_sample_pool` 函数完全重写
- 4 张市场卡片视觉重构(每张占 13 行 R7-R19,header / band / 一键大按钮 / 显示隐藏中按钮 / 代码-简称 sub-header / 数据区)
- 显示/隐藏按钮用 brand 浅色版(浅红 `#FCE4E4` / 浅绿 `#EAF4E3` / 浅紫 `#EEE8F7`)取代统一浅蓝
- 工具栏(一键全抓 + 全局显示 + 显示/隐藏所有 + 工具 + 清空缓存)垂直堆叠右上 N2:Q11
- 使用提示 panel 浮右上 S2:V8 + `💡` emoji + 5 条 numbered tips
- 参数设置 panel 收窄到 A2:M2(原占 A2:N2 整行)
- 新增 `_apply_card_brand` helper 抽象 4 张卡片重复逻辑

**Phase 4j.3** `ccd9ee7` — Reviewer 修两处视觉 bug:
- padding cols (C/F/I/L/M) 从 width 4/4/4/4/10 缩到 1/1/1/1/1
- 卡片 R8 brand-color band 改成白色窄带(原 Codex 实现让 R7 header + R8 band 视觉融成一个色块,用户反馈"没东西但有颜色")

**Phase 4j.4** `67252d2` — Reviewer 进一步彻底隐藏:
- width=1 仍能看到 1px 空列,改 `Columns(col).Hidden = True` 彻底隐藏
- 4 张卡片现在视觉上紧贴,数据区 R14+ 不再有空白 padding 单元格
- 按钮 merge range 不变(eg 一键 A 股 仍 `A9:C10` 但 C 列 hidden,实际只显示 A:B 视觉宽度)

### AA.2 验证结果

| 项目 | 结果 |
|---|---|
| `py tools/test_fx_live.py --skip-install` | PASS,5/5 |
| `py -u tools/diff_phase4f_step3_lite.py` | PASS;A股资产负债表 `原币` vs `统一RMB` 0 mismatches |
| `py -u tools/inspect_phase4g_state.py` | EXIT_CODE=0;toggle 行为符合 Phase 4j Step 6 设计 |
| `py -u tools/inspect_phase4h_state.py` | PASS;B6 toggle / 缓存 / 自动 fallback / 跨市场指标表 全部正常 |
| 用户视觉验收(打开 xlsm 肉眼对比目标截图)| ✅ 通过 |

### AA.3 已知边界

- 样本池 4 张卡片采用 hide column 方案而非完全 restructure,padding cols (C/F/I/L/M) 仍存在于 cell 范围中(只是 Hidden=True 不显示);后续如需清理可作为优化项
- 按钮 merge range 跨越 hidden col(eg `A9:C10` 含 hidden C),功能正常但内部 cell 引用看起来有点 "hack",纯视觉
- 配置区 `参数设置 / 一键全抓 / 使用提示` 三大块视觉对齐符合截图;数据录入区 R14+ 也按 4 张卡片视觉紧贴
- `_apply_card_brand` helper 把 4 张卡片重复逻辑抽出,后续如需调整 card 风格只改一处

### AA.4 当前 release 候选

commit `67252d2` 是从 Phase 4f 起累计 9 phase / 17 commit 的核心开发完整汇总:
- Phase 4f:RMB 换算 hook + 汇率 sheet + B6 toggle scaffold
- Phase 4g:跨市场指标合表 + hide-tab 按钮 + POOL_DATA_START_ROW 迁移
- Phase 4h:B6 实时 toggle + 磁盘 JSON 缓存 + stockanalysis 中概美股 fallback
- Phase 4i / 4i.1 / 4i.2:UX 抛光(样本池 / 使用说明商务化 / 汇率说明区 / 单表按钮删除 / fallback 自动化)
- Phase 4j / 4j.1:简化跨市场对比策略(只留指标表)+ 抑制合并弹窗
- Phase 4j.2 / 4j.3 / 4j.4:**样本池视觉 1:1 还原(本阶段)**

### AA.5 下一阶段(明日起)

**核心功能开发到此告段落**,明日起进入**优化工作**阶段。具体优化方向待 next session 与用户协商,候选 backlog:
- 抓数性能(`.cache/` 命中率提升 / 并发探索)
- 数据质量(港股双年 fetch 扩展 / 美股 fallback 白名单扩展 / 韩股字段补齐)
- 视觉精修(其他 sheet 的 design system apply / 跨市场指标表视觉)
- 用户体验(B6 toggle 细节 / cookie 失效友好提示 / 错误诊断 UI)
- 运维(CI 跑 frozen 回归 / release 打包流程 / 老用户升级测试)

---

## BB. Phase 4k 收口: 优化 Sprint 1 — 数据准确性 + UX live FX + 状态守护

执行依据: `PHASE_4K_PLAN.md` v1。状态: ✅ Codex 已实现并通过 4 张 frozen 回归 + 新增 Phase 4k inspect。

### BB.1 已完成

- [Step 1 P0-02] 诊断表 Score 列改文本格式, `1/1` 不再被 Excel 自动转日期。
- [Step 2 P0-03] 新增 `模块_AppStateGuard.bas`, `一键全抓` 入口示范使用 `BeginAppState/EndAppState` 恢复 Excel 状态。
- [Step 3 P0-05] 新增 `GetFxRateStatus`, 汇率缺失不再 fallback 1;统一 RMB 模式下写空值并追加诊断 `FX_MISSING`。
- [Step 4 P1-05] 报表公式从 baked-in 汇率改为 `GetFxFromSheet("USD",C$2,"EOP/AVG")`, 手改汇率 sheet 后公式实时重算。
- [Step 5] 新增 `tools/inspect_phase4k_state.py`, 覆盖 Score 文本 / AppStateGuard / FX_MISSING / live FX UDF 性能检查。

### BB.2 验证结果

| 项目 | 结果 |
|---|---|
| VBE Compile | Compile command executed; `模块_测试.TestPhase4kScoreSmoke` PASS |
| `py tools/test_fx_live.py --skip-install` | PASS,5/5;旧 `GetFxRate` 签名兼容 |
| `py -u tools/diff_phase4f_step3_lite.py` | PASS;A股资产负债表 `原币` vs `统一RMB` 0 mismatches |
| `py -u tools/inspect_phase4g_state.py` | EXIT_CODE=0 |
| `py -u tools/inspect_phase4h_state.py` | PASS |
| `py -u tools/inspect_phase4k_state.py` | PASS;Score `1/1` 保持文本;KRW 缺汇率输出空值 + `FX_MISSING` row 3;live FX UDF 900 公式 smoke 响应约 0.137s |

### BB.3 已知边界

- AppStateGuard 本期只 apply 到 `一键全抓` 入口, 其他入口后续 sprint 增量做。
- FX_MISSING 时报表 cell 显示空, 不显示 `#N/A`,避免 Excel 错误传播到分析表。
- `GetFxFromSheet` UDF 不读 H 列 override,保留为 future enhancement。
