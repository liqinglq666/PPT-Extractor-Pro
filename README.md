# 📹 PPT-Extractor-Pro：全学科通用的“科研素材收割机”

![Python](https://img.shields.io/badge/Python-3.9%2B-007EC6)
![Computer Vision](https://img.shields.io/badge/Tech-OpenCV_Target_Tracking-orange)
![Methodology](https://img.shields.io/badge/Method-Spatio_Temporal_Saliency-green)
![Productivity](https://img.shields.io/badge/Goal-Research_Material_Collection-purple)
![License](https://img.shields.io/badge/License-MIT-lightgrey)

## 👋 起因 (Why I made this)

各位被组会、学术会议、网课回放折磨的**科研特种兵**们，Hello！👋

众所周知，科研的起步往往是从**模仿**开始的。
看到行业大牛在顶会上的 PPT 做得那么漂亮，逻辑那么丝滑，图表那么精致，我们第一反应是什么？
**——“截下来！以后写论文/做汇报时参考（抄）一下！”**

但是，手动截图真的太痛苦了：
1.  **废手**：一场 2 小时的讲座，全程紧盯屏幕按 `PrtSc`，右手食指差点腱鞘炎。
2.  **废眼**：稍不留神低头回个消息，大牛翻页了，错过了最关键的数据图。
3.  **废脑**：截下来的图全是散的，想复盘大牛的“讲故事逻辑”，还得自己在文件夹里拼图，心态崩了。

为了**高效建立自己的科研素材库**，也为了**从大牛的 PPT 里汲取灵感**，我搓了这个工具。
**核心宗旨：视频丢进去，我去拿外卖，回来直接看 PDF。把别人的精华，变成我的素材。**

---

### 🧬 核心算法架构 (Core Algorithmic Architecture)



本系统摒弃了低效的“时钟驱动（Timer-driven）”采样策略，转而采用了一套 **时空联合的视觉显著性感知框架 (Spatio-Temporal Visual Saliency Perception Framework)**。

#### 1. 空间域注意力屏蔽机制 (Spatial Domain Attentional Masking)
针对非结构化视频流中的高频噪声（如讲师肢体语言、UI 控件闪烁、弹幕），我引入了 **ROI 空间滤波**技术。
通过构建一个正交的感兴趣区域（Region of Interest），系统能够对视频帧进行 **语义分割** 的预处理，将信噪比（SNR）极低的背景区域在硬编码层级直接剔除。
> **(人话：只盯着 PPT 那块地，讲师的脸、底下的进度条，通通不看。)**

#### 2. 时域动能收敛检测 (Temporal Kinetic Convergence Detection)
为了精准捕获 PPT 的“稳态”，算法计算连续帧之间的 **均方误差 (MSE)** 及其梯度变化。
我将视频流视为一个动态的时间序列，只有当帧间差异的 L2 范数收敛至一个极小的 $\epsilon$ 邻域，并保持 $N$ 个时间步长（Time Steps）的**视觉静默**(Visual Quiescence)时，状态机才会触发捕获中断。
> **(人话：等翻页动画彻底停稳了再动手，保证抓下来的图是高清无重影的，方便后续 OCR 识别文字。)**

#### 3. 感知哈希拓扑去重 (Perceptual Hashing Deduplication)
为了解决视觉冗余问题，系统维护了一个 **特征指纹注册表 (Feature Fingerprint Registry)**。
每一帧候选图像都会经过 **dHash (差分哈希)** 算法降维，映射为一个 64-bit 的特征向量。通过计算当前帧与历史帧向量之间的 **汉明距离 (Hamming Distance)**，我们能够以 $O(1)$ 的时间复杂度判定语义重复性。
> **(人话：如果大牛在一页 PPT 上讲了 20 分钟没翻页，系统绝对不会给你存两张一样的图。素材库要精简，不要垃圾。)**

---

## 🚀 几个极其舒适的功能 (Features)

### 1. 🎯 "时空双域" 精准打击 (Spatio-Temporal Precision)
我们不只截取屏幕，我们是在做**外科手术级的素材提取**。
* **📐 空间域 - ROI 狙击模式**：
    * 点击 `[+] 设定 ROI`，直接在画面上画框。
    * **效果**：像切掉面包边一样，把视频四周的黑边、弹幕、右下角讲师的头像通通切掉，只留纯净的 PPT 内容。
* **✂️ 时域 - 可视化时间轴裁剪**：
    * **痛点**：大牛讲座前 10 分钟是废话，最后 20 分钟是无聊的 Q&A？
    * **解决**：内置可视化剪辑器，支持**帧级微调**（上一帧/下一帧）。你可以精准锁定大牛开始讲干货的那一秒，只处理精华片段，绝不浪费算力。

### 2. 📂 强迫症福音的文件治理 (Structured Asset Management)
拒绝桌面上一堆 `截图1.png`, `截图2.png` 这种让科研人崩溃的垃圾堆！
* **🗂️ 项目制管理**：
    * 输入项目名（如 `Prof_Li_CVPR2025`），软件自动清洗非法字符并创建独立目录。
* **📦 自动归档与合成**：
    * `Runs/`: 存放所有抽取出的高清原始图片（适合做论文插图素材）。
    * `PDFs/`: 程序跑完**自动**调用内置引擎，将图片合并为高清 PDF。
    * **贴心细节**：任务结束时**自动弹出文件夹**，无需你去硬盘里翻箱倒柜。

### 3. ⚡ 性能与交互的平衡 (Performance & Interaction)
* **📺 摸鱼开关 (LIVE Toggle)**：
    * **ON (调试模式)**：右上角实时播放截取过程，看着 PPT 一张张被抓取，极其解压。
    * **OFF (狂暴模式)**：一旦确认参数无误，关掉它！系统将跳过 UI 渲染，**全速运行**（速度提升 30%+）。此时你可以把窗口最小化，假装自己在读文献。
* **🎛️ 参数热调节**：
    * **灵敏度 (Threshold)**：PPT 背景太花容易误触？调高它。
    * **防抖 (Stability)**：翻页动画太慢？调高防抖等级。
    * **所见即所得**：拖动滑块实时生效，无需重启程序，比调教你的实验设备容易多了。

---

## 📸 界面长这样

*( 界面：左边配置参数，右边实时监控，底部状态反馈 )*

<div align="center">
  <h3>🖥️ 主界面 (Main UI)</h3>
  <img src="./assets/img_1.png" width="100%" alt="Main Interface" />
  
  <br><br>
  
  <h3>⚙️ 参数配置 (Settings)</h3>
  <img src="./assets/img_3.png" width="100%" alt="Settings Interface" />
</div>

---

## 💾 怎么下载 (Download)

考虑到大家的电脑环境五花八门（有的还在用 Win7，有的 Python 环境乱得像盘丝洞），我打包了个 **免安装的 .exe 版本**。
**下载 -> 解压 -> 双击** 就能跑。

* 🚀 **下载地址**：👉 **[点击这里下载最新版本 (Latest Release)](https://github.com/liqinglq666/PPT_Extractor_Pro/releases/latest)**
* **系统支持**：Win 10 / Win 11 

---

## 📂 喂饭级操作指南 (User Guide)

**⚠️ 简单三步，把视频变成科研素材：**

### 1. 准备视频 (Prepare)
把你要拆解的视频准备好。支持 `.mp4`, `.avi`, `.mkv`。建议画质高点，方便后续作为高质量素材引用。

### 2. 配置参数 (Configure)
* **设定 ROI (重点!)**：点击 `[+] 设定 ROI 扫描区域`，框它！
    * *小技巧*：框选时尽量避开字幕区，不然字幕一变，程序就以为 PPT 翻页了。
* **项目命名**：给这次任务起个名，比如 `AI_Trend_Report`。
* **参数微调**：
    * **判定阈值**：默认 10。如果 PPT 背景很花（渐变色、大图背景），调高点（12-15）；如果是纯白底学术风，默认就行。
    * **防抖等级**：默认 5。如果视频里 PPT 翻页动画慢得像老太太过马路，建议调高到 8-10。

### 3. 启动 (Run)
点击 **INITIALIZE ENGINE**。
* 此刻你可以去**冲咖啡 / 拿外卖 / 刷抖音**。
* 跑完后，去 `输出目录/项目名/PDFs` 里收货。

---

## 🌍 适用场景 (Use Cases)

1.  **模仿大牛逻辑**：下载顶会/大牛讲座视频，提取 PPT，研究他们是如何**排版**、如何**讲故事**、如何**设计图表**的。
2.  **构建素材库**：把视频里精彩的架构图、数据图提取出来，分类保存，下次做汇报没灵感时翻一翻。
3.  **网课/MOOC 笔记**：Coursera, B站, 学堂在线……只想要讲义，不想要废话。
4.  **学术会议记录**：Zoom/腾讯会议录屏，提取汇报人的 Slide，方便组会讨论。

---

## 🛠️ 开发者指南 (For Developers)

如果你想魔改算法或者加个 OCR 自动提取文字功能，欢迎 Clone。

本项目基于 Python 3.9+ 构建，主要依赖：`tkinter` (UI), `ttkbootstrap` (美化), `opencv-python` (视觉核心), `pillow` (图像处理)。

### 快速开始

```bash
# 1. 克隆仓库
git clone https://github.com/liqinglq666/PPT-Extractor-Pro.git

# 2. 安装依赖 
pip install -r requirements.txt

# 3. 启动
python main.py