@echo off
:start
cls
echo 1.请拖入(输入)图片文件夹(路径)
echo 2.^^_^^ 然后按回车键确认 ^^_^^

set /p "image_dir="
:: 用户按回车后清屏
cls

set "secret_id=您的SecretId"
set "secret_key=您的SecretKey"

python tencent_table_ocr_batch.py ^
  --image_dir %image_dir% ^
  --output_dir ./results ^
  --log_dir ./logs ^
  --secret_id "%secret_id%" ^
  --secret_key "%secret_key%"

echo.
echo 按任意键继续处理下一个文件夹...
pause >nul
goto start