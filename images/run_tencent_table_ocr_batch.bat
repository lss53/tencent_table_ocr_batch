@echo off
:start
cls
echo 1.������(����)ͼƬ�ļ���(·��)
echo 2.^^_^^ Ȼ�󰴻س���ȷ�� ^^_^^

set /p "image_dir="
:: �û����س�������
cls

set "secret_id=����SecretId"
set "secret_key=����SecretKey"

python tencent_table_ocr_batch.py ^
  --image_dir %image_dir% ^
  --output_dir ./results ^
  --log_dir ./logs ^
  --secret_id "%secret_id%" ^
  --secret_key "%secret_key%"

echo.
echo �����������������һ���ļ���...
pause >nul
goto start