# bi_skill.py 路径解析

`config.py` 中 `_resolve_bi_skill_path()` 按下列顺序查找 `bi_skill.py`，命中即停：

1. **环境变量 `BI_SKILL_PATH`**（推荐，跨电脑最稳）
2. `~/.workbuddy/skills/bi_skill/bi_skill.py`（标准安装位置）
3. 与本 Skill 同级的 `../bi_skill/bi_skill.py`
4. `../.workbuddy/skills/bi_skill/bi_skill.py`

四个候选都不命中时，仍返回候选 1 的路径，交给后续步骤报"文件不存在"，便于排查。

## 切换路径

PowerShell：
```powershell
$env:BI_SKILL_PATH = "D:\tools\bi_skill\bi_skill.py"
```

Bash / WSL：
```bash
export BI_SKILL_PATH="$HOME/tools/bi_skill/bi_skill.py"
```

要持久生效，写进系统环境变量或 shell 配置文件（`.bashrc` / PowerShell `$PROFILE`）。

## 验证

```bash
python -c "from config import BI_SKILL_PATH; print(BI_SKILL_PATH, BI_SKILL_PATH.exists())"
```

预期输出：实际路径 + `True`。`False` 即说明候选都未命中，按上面"切换路径"显式注入。
