# Windows构建与GitHub交付

仅在当前任务或会话已授权提交/发布的范围内执行。准备和验证代码不需额外审批；已有发布授权时直接完成。

## 发布前的关键检查

读取实际`.github/workflows/windows-release.yml`、`app.py`、README与CHANGELOG，不凭此文件假定流水线已支持新版本。

**v2.0流水线存在硬编码版本**：EXE名、ZIP名、artifact名、tag和`gh release upload --clobber`均指向v2.0，且已存在tag不会移动。后续新版本不能直接重跑它，否则会把新二进制覆盖到仍指向旧源码的v2.0发布中。

发布v2.1等新版本时应在本次代码改动中先解决版本一致性：统一版本来源或新增版本输入，并同步应用标题、自检断言、EXE名、ZIP名、artifact、release tag、说明及下载链接。可以按当前项目复杂度选择最小可靠实现，不必为了文档强行重构。

- 新tag对应本次已测提交；若已存在，应确认它确实是预期提交，不能静默覆盖不同源码版本的资产。
- 默认新版本创建新tag并保留旧发布。只有用户明确要求修订同一版本且提交/资产一致性已经处理时，才替换资产。
- 记录Python、应用依赖、PyInstaller和Windows环境版本，构建包不携带原工具、真实数据、配置密码或临时日志。

## 验证与构建

1. 本地相关测试与必要回归通过，提交代码和文档，核对远端目标提交。
2. 检查GitHub凭据确实有仓库和Actions所需权限。可用连接工具优先；若缺少调度/上传能力，使用已认证的`gh`，不输出令牌。
3. 在Windows构建。v2.0使用Python3.12和PyInstaller6.20.0，示例命令中的版本名按本次目标替换：

```sh
python -m pip install -r requirements.txt pytest pyinstaller==6.20.0
python -m pytest -q
python -m PyInstaller --noconfirm --clean --onefile --windowed --name SalaryConsolidator-vNEXT --collect-all msoffcrypto app.py
```

4. 实际启动打包EXE，等待退出并读取`--self-test`结果。设置超时，失败则查看资源/导入/Tk日志并修复；只看到dist文件不算验证通过。
5. 本次变更涉及额外资源、解密、xls或RAR后端时，增加相应打包后用例。Windows测试不能被Linux源码测试替代。
6. ZIP包含实际EXE及适用使用说明，计算SHA256，上传对应新版本Release并保留构建run/commit关联。

## Actions操作模板

先确认工作流已经完成版本参数化或已改为本次正确版本；以下不是可盲目粘贴的发布命令：

```sh
gh workflow run windows-release.yml --repo ZY-2012/tmp_project --ref <待发布分支或提交>
gh run list --repo ZY-2012/tmp_project --limit 5
gh run view <本次run-id> --repo ZY-2012/tmp_project
gh run view <失败run-id> --repo ZY-2012/tmp_project --log-failed
```

锁定本次run的head SHA，不把别人的绿色run或旧构建当作本次结果。脚本/依赖失败则修复后对新提交重新调度；重跑旧run不会包含修复。

正常等待构建时定期报告实质状态，不结束任务让用户自行等待，也不长时间阻塞所有沟通。当前会话没有外部发布能力时，明确已完成的代码/测试与具体阻塞，不编造下载链接。

## 回下载验收

发布成功后至少：

- 读取Release元数据，确认非draft、tag与目标源码对应、资产名称和大小正确。
- 下载公开包及校验文件到临时目录，计算SHA256并比较。
- 用zipfile检查成员CRC、目录与必要EXE；不要把“上传成功”当成文件完整性校验。
- 核对包内说明及版本，最终给出实际可访问下载链接、源码/对照链接和验证限制。

## 仅提交文档或skill

不需要重打应用、不需要改变应用版本/tag，也不需要跑整套工资业务测试。验证内容与链接后提交相应文件；若交付skill压缩包，可追加独立命名的skill资产，不能覆盖应用ZIP或应用校验文件。说明skill文档基线与应用版本，避免让人误以为应用二进制已更新。

仓库另有`.github/workflows/skill-package.yml`可单独验证并打包本技能，将ZIP及校验值附加到指定已存在Release，不重构建应用、不移动应用tag。后续修订使用新的package_name（如附带技能版本号），流程默认不覆盖同名资产。若直接上传通道无法认证，可使用已授权的Actions身份运行此技能专用流程，不必索取用户令牌。
