# Windows 电池深度分析工具

一键生成中文 HTML 电池健康分析报告，读取 Windows `powercfg /batteryreport` 数据，解析电池信息、使用历史、容量变化趋势，并生成可视化图表。

## 功能

- 自动调用 `powercfg /batteryreport` 生成原始报告，或直接读取已有 HTML 文件
- 解析多块电池的设计容量、满充容量、循环次数，计算健康度与损耗
- 支持读取 Win32_Battery 实时状态（充电/放电功率、电流、电压）
- 纯 CSS/HTML 图表，无需额外渲染依赖：
  - 电池健康度对比（横向条形图）
  - 各电池容量对比（设计 vs 满充）
  - 电池容量变化历史（折线图 + 中值趋势线）
  - 续航估算历史（折线图）
  - 月度使用时长历史（堆叠柱状图）
  - 最近使用活动分析（放电事件 + 状态占比）
- 长周期数据自动按月聚合
- 所有图表自适应容器宽度，支持移动端查看

## 环境要求

- Windows 10/11（依赖 `powercfg` 命令）
- Python 3.8+

## 安装

```bash
git clone https://github.com/<你的用户名>/battery-report.git
cd battery-report
pip install lxml
```

## 使用方法

### 方式一：自动生成（推荐）

直接运行，程序会自动调用 `powercfg /batteryreport` 生成报告并分析：

```bash
python battery_report.py
```

### 方式二：读取已有报告

如果你已经通过 `powercfg /batteryreport` 生成了 HTML 文件，可以直接传入：

```bash
python battery_report.py "C:\path\to\battery-report.html"
```

### 命令行参数

| 参数 | 说明 |
|------|------|
| `input`（可选） | 已有的 battery-report.html 路径，不填则自动生成 |
| `--no-open` | 生成后不自动打开浏览器 |

## 输出

报告生成在 `%TEMP%\BatteryReport\笔记本电池深度报告.html`，并自动用默认浏览器打开。

报告内容包括：

- **系统与报告信息**：电脑型号、BIOS、Windows 版本、当前电量、充电/放电功率等
- **电池概览**：设计容量合计、满充容量合计、组合健康度、循环次数
- **单电池详情**：健康度环形图、制造商、序列号、化学类型、每周期损耗
- **当前续航估算**：按设计容量与满充容量分别估算
- **可视化分析**：6 张图表
- **综合评估与建议**：基于健康度和循环次数的自动建议

## 健康度计算

```
健康度 = 满充容量 ÷ 设计容量 × 100%
```

| 健康度 | 状态 |
|--------|------|
| ≥ 80% | 健康 |
| 60%–80% | 衰减 |
| < 60% | 严重衰减 |

## 项目结构

```
battery-report/
├── battery_report.py          # 主程序
├── README.md                  # 说明文档
└── 笔记本电池深度报告.html      # 示例输出报告
```

## 依赖

| 依赖 | 用途 |
|------|------|
| [lxml](https://lxml.de/) | 解析 Windows battery-report HTML |

所有图表使用纯 CSS/SVG 渲染，无需 matplotlib 或其他绘图库。

## 许可证

MIT License
