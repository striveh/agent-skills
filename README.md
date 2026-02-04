# agent-skills

## icp-batch-skill
- **作用**：批量处理域名备案查询。对 `domains.xlsx` 中表头“链接”列提取域名，优先用本地缓存 `icp_results.csv`，缺失或失败时调用 `https://domainicp.market.alicloudapi.com/do?domain=...`（AppCode 鉴权），生成成功清单 `icp_success.csv`，并将备案主体/备案号回填至 Excel 末尾两列。
- **使用**：
  ```bash
  cd skills/icp-batch-skill
  export APP_CODE=<你的AppCode>
  python scripts/run_icp_batch.py \
    --workbook domains.xlsx \
    --cache icp_results.csv \
    --success icp_success.csv
  ```
  可选参数：`--host`、`--path`（默认 `https://domainicp.market.alicloudapi.com`、`/do`），`--sleep` 控制调用间隔。默认文件名与表头均可覆盖。依赖：Python3、`openpyxl`、`requests`。

- **Windows EXE（非技术人员）**：GitHub Actions 会生成可执行文件 `icp-batch-skill.exe`（Artifacts 下载）。
  - 双击运行后会提示输入 AppCode，并弹出文件选择器选择要处理的 Excel；进度与错误会用弹窗提示。
  - 可选：同目录放 `appcode.txt`（一行 AppCode）自动读取；输出文件默认保存在所选 Excel 同目录。
