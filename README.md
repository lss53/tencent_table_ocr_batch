### 使用说明文档

**腾讯云 OCR 表格批量识别工具 v1.0**  
*（支持 PNG/JPG/BMP 等格式图片，自动生成 Excel 表格）*

---

#### 一、准备工作
1. **安装 Python**  
   - 下载 Python 3.7+：[Python 官网](https://www.python.org/downloads/)
   - 安装时勾选 `Add Python to PATH`

2. **安装依赖库**（CMD 执行）：
```bash
pip install tencentcloud-sdk-python openpyxl
```

3. **获取腾讯云密钥**  
   - 访问腾讯云官网：https://cloud.tencent.com。若未注册，点击右上角「注册」，使用手机号、邮箱或微信/QQ等方式完成注册；若已有账号，直接点击「登录」。
   - 进入https://console.cloud.tencent.com/developer/auth，选择「个人实名认证」，按提示上传身份证信息并完成验证。  
   - 进入腾讯云控制台：https://console.cloud.tencent.com，在顶部导航栏选择「产品」> 搜索「文字识别」，进入OCR控制台后，系统会提示“立即开通”。点击「开通」，根据提示确认服务协议，完成开通。
   - 进入https://console.cloud.tencent.com/cam/capi。创建 API 密钥（保存 `SecretId` 和 `SecretKey`）。  
     强烈建议您使用`子账号密钥`！
     - 用户列表 > 新建用户 > 快捷创建 > 输入`用户名` > 用户权限: 搜索ocr，勾选`QcloudOCRReadSelfUinUsage`并确定 > 创建用户
     - 点上面创建的`用户名` > API密钥 > 新建密钥

---

#### 二、文件配置
1. **创建项目文件夹**（示例结构）：
```
TableOCR_Tool/
├── tencent_table_ocr.py   # Python 主程序
├── run_table_ocr.bat      # 批处理脚本
├── logs/                  # 日志目录（自动创建）
└── results/               # 结果目录（自动创建）
```

2. **编辑批处理脚本**（用记事本打开 `run_table_ocr.bat`）：  
   修改第 16-17 行为您的密钥：
```bat
set "secret_id=您的SecretId"
set "secret_key=您的SecretKey"
```

---

#### 三、使用步骤
1. **准备图片**  
   - 将包含表格的图片放入任意文件夹（支持 PNG/JPG/JPEG/BMP/GIF/TIFF）
   - 单张图片需 ≤3MB
   - 本工具仅适用于[2024年重庆市普通高校招生信息表本科批-物理-平行志愿](https://www.cqksy.cn/web/article/2024-07/21/content_6230.html)的图片表格，可用[图片助手(ImageAssistant)](https://www.pullywood.com/ImageAssistant/)批量下载。
   
2. **运行识别工具**：
   
   - 双击运行 `run_table_ocr.bat`
   - 按提示 **拖入图片文件夹** 到窗口（或直接输入路径）
   ```bat
   1.请拖入或输入图片文件夹路径
   2.^^_^^ 然后按回车键确认 ^^_^^
   [拖入文件夹到此]
   ```
   
3. **查看结果**：
   
   - Excel 文件生成在 `results/` 目录
   - 日志文件保存在 `logs/` 目录
   - 失败图片列表保存在 `results/失败报告_时间戳.txt`

---

#### 四、高级选项（可选）
通过命令行直接运行：
```bash
python tencent_table_ocr.py \
  --image_dir "图片路径" \
  --output_dir "输出目录" \
  --secret_id "您的SecretId" \
  --secret_key "您的SecretKey" \
  --max_workers 4         # 并发数(默认2)
  --batch_size 20         # 分批保存数量(默认10)
```

---

#### 五、注意事项
1. 首次运行可能触发 Windows 防火墙警告，需允许访问
2. 每 10 张图片自动保存临时结果（防中断）
3. 特殊表格建议：
   - 避免复杂合并单元格
   - 确保图片分辨率 > 300dpi
   - 深色背景建议转为黑白
