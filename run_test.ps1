$env:PYTHONIOENCODING="utf-8"
$url = (python get_url.py).Trim()
Set-Clipboard -Value $url
python clip_save.py
