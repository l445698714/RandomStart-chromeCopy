<div align="center">

# Chrome 分身启动器

[![Python](https://img.shields.io/badge/Python-3.x%2B-3776AB.svg?style=flat&logo=python&logoColor=white)](https://www.python.org/)
[![Windows](https://img.shields.io/badge/Windows-10%2B-0078D6.svg?style=flat&logo=windows&logoColor=white)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**作者：l445698714**
<br />
[![GitHub](https://img.shields.io/badge/GitHub-l445698714-lightgrey.svg?style=flat&logo=github&logoColor=black)](https://github.com/l445698714)
[![X (Twitter)](https://img.shields.io/badge/X_StayrealLoL-1DA1F2.svg?style=flat&logo=x&logoColor=white)](https://x.com/StayrealLoL)

</div>

## ⚠️ 免责声明

1.  本工具按“原样”提供，不作任何明示或暗示的保证。
2.  用户应对使用本工具的任何行为及其后果负全部责任。
3.  开发者不对因使用（或滥用）此工具而造成的任何直接或间接损害负责。
4.  请确保您的使用行为符合您所在地区及相关平台的法律法规。

## ❇️ 工具介绍
`Chrome 分身启动器` 是一个基于 PyQt5 的 Windows 桌面应用程序，旨在帮助用户高效地批量管理和启动多个 Chrome 浏览器分身。用户可以方便地指定启动范围、数量，设置启动延迟，并能快速关闭所有或指定范围的分身。程序会记录已启动的分身编号，避免重复操作，并提供清晰的状态反馈。

## 💖 支持项目
如果您觉得这个工具对您有帮助，可以请我喝杯蜜雪冰城，增加更多功能和分享更多小工具：
<br />
<img src="https://github.com/user-attachments/assets/b7810000-78d3-4c6b-a10d-cee4d22d6845" alt="Donation QR Code 1" width="200"/>
<img src="https://github.com/user-attachments/assets/11952997-dd5d-4311-a085-8145acdb4950" alt="Donation QR Code 2" width="200"/>

## ❇️ 功能特性

-   **多种启动模式**：
    -   随机启动：在指定范围内随机选择并启动一定数量的 Chrome 分身。
    -   指定范围启动：启动用户定义范围内的所有有效 Chrome 分身。
    -   依次启动：从指定编号开始，按顺序逐个启动 Chrome 分身。
-   **分身管理**：
    -   自定义快捷方式文件夹：允许用户指定存放 Chrome 分身快捷方式 (`.lnk` 文件) 的目录。
    -   关闭所有 Chrome 窗口：一键快速关闭所有侦测到的 Chrome 浏览器进程。
    -   指定范围关闭：关闭指定编号范围内的 Chrome 分身进程。
    -   在已打开的分身中打开网址：扫描当前运行的 Chrome 实例，并在所有独立实例中打开指定 URL。
-   **用户体验**：
    -   实时状态显示：在界面上清晰展示当前操作状态、已启动编号列表、范围内剩余未启动数量等信息。
    -   可调节启动延迟：用户可以自定义每次启动操作之间的间隔时间（秒），以防止系统因快速连续操作而过载。
    -   设置持久化：程序会自动保存用户的常用设置（如快捷方式路径、范围、延迟等）。
    -   内置亮色主题，界面友好。

## ❇️ 环境要求

-   **操作系统**: Windows 10 或更高版本。
-   **Python 版本**: Python 3.x (建议 3.7+)。
-   **必要 Python 库**:
    -   `PyQt5` (用于图形用户界面)
    -   `psutil` (用于进程管理)

## ❇️ 运行教程

1.  **克隆或下载项目**
    *   您可以从本 GitHub 仓库 (`https://github.com/l445698714/auto-chrome`) 下载最新的源代码。如果使用 Git，可以克隆仓库：
        ```bash
        git clone https://github.com/l445698714/auto-chrome.git
        cd auto-chrome
        ```

2.  **安装 Python 环境**
    *   确保您的 Windows 系统已安装 Python 3.x。您可以从 [Python官网](https://www.python.org/downloads/) 下载并安装。
    *   安装 Python 时，建议勾选 "Add Python to PATH" (将 Python 添加到环境变量)。

3.  **安装项目依赖**
    *   打开命令行工具 (如 Command Prompt 或 PowerShell)，进入项目根目录 (即 `auto-chrome` 文件夹)。
    *   运行以下命令安装所需的 Python 库：
        ```bash
        pip install PyQt5 psutil
        ```
    *   如果下载速度较慢，可以尝试使用国内镜像源：
        ```bash
        pip install PyQt5 psutil -i https://pypi.tuna.tsinghua.edu.cn/simple
        ```

4.  **运行程序**
    *   在项目根目录下，通过以下命令运行主程序脚本：
        ```bash
        python Chrome_launcher.py
        ```
    *   如果脚本目录中包含 `ico.ico` 文件，程序启动时会尝试加载它作为窗口图标。

## ❇️ 使用说明

1.  **首次配置**：
    *   启动程序后，首先点击"分身路径设置"区域的"浏览..."按钮，选择包含您的 Chrome 分身快捷方式（`.lnk` 文件）的文件夹。这些快捷方式应以数字命名（例如 `1.lnk`, `2.lnk`）。
2.  **参数设置**：
    *   **启动范围设置**: 输入希望操作的分身编号的起始和结束值（例如 `1` 到 `100`）。
    *   **启动数量设置**: （主要用于随机启动和首次依次启动）输入希望一次启动的分身数量。
    *   **指定范围设置**: 输入一个明确的范围（例如 `1-10` 或 `5-20`），用于"指定启动"、"依次启动"和"指定关闭"功能。
    *   **启动延迟设置**: 设置每次启动或关闭操作之间的延迟时间（秒），以避免系统卡顿。
    *   **打开网址**: 输入希望在已运行分身中打开的完整 URL。
3.  **执行操作**：
    *   根据需求点击对应的功能按钮："随机启动"、"指定启动"、"依次启动"、"在已打开的分身中打开网址"、"指定关闭"、"关闭所有"。
4.  **查看状态**：
    *   程序主界面下方的状态区域会显示详细的操作反馈和当前已记录的已启动分身列表。
    *   程序底部的状态栏也会提供简洁的状态摘要。

## ❇️ 注意事项

-   **快捷方式命名**：请确保您在"快捷方式目录"中指定的文件夹内，包含了以数字命名的 Chrome 分身快捷方式（例如 `1.lnk`, `2.lnk`, ..., `100.lnk`）。程序依赖这些数字文件名来识别和操作分身。
-   **管理员权限**：根据系统设置和 Chrome 安装位置，某些操作（特别是关闭进程）可能需要程序以管理员权限运行才能成功。
-   **状态记忆**：程序会通过设置文件 (`QSettings`) 记住已启动的分身编号以及其他界面设置，直到程序关闭或重新启动。在程序启动时，它会尝试与当前运行的 Chrome 进程进行同步，以更新记录的已启动分身列表。
-   **启动延迟**：合理设置"启动延迟"可以有效防止因快速连续启动过多 Chrome 实例而导致系统响应缓慢或卡顿。
-   **关闭所有功能**： "关闭所有"按钮会尝试终止所有名为 `chrome.exe` 的进程。请谨慎使用，确保没有重要的未保存工作。


## 许可证

本 `Chrome 分身启动器` 项目采用 MIT 许可证。详情请参阅仓库中的 `LICENSE` 文件（如果尚未创建，默认为 MIT 条款）。

