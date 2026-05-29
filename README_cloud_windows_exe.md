# Build Windows EXE in GitHub Actions

这个项目已经配置好 GitHub Actions，可以在云端生成 Windows 用户可直接双击运行的 `基金净值分析工具.exe`。

## 使用步骤

1. 新建一个 GitHub 仓库。
2. 上传这些文件：
   - `fund_analyzer_gui.py`
   - `fund_analyzer_windows.spec`
   - `requirements-windows.txt`
   - `.github/workflows/build-windows-exe.yml`
3. 进入 GitHub 仓库的 `Actions` 页面。
4. 选择 `Build Windows EXE`。
5. 点击 `Run workflow`。
6. 等任务完成后，在任务页面下载 artifact：`windows-exe-基金净值分析工具`。
7. 解压后得到 `基金净值分析工具.exe`。

## 重要说明

- 使用这个 exe 的 Windows 电脑不需要安装 Python。
- GitHub 云端构建时会临时安装 Python 和依赖，但这是在 GitHub 的 Windows 机器上完成，不是在用户电脑上。
- 如果 Windows 提示未知发布者，是因为没有购买代码签名证书；选择“更多信息”后继续运行即可。
