# Resumes Screener

Flask + Waitress web service for screening resume documents against a job description.

## Run locally

```powershell
python -m pip install -r requirements.txt
$env:RESUMES_SCREENER_HOST = "0.0.0.0"
$env:RESUMES_SCREENER_PORT = "80"
python .\resumes_screener.py
```

Open `http://<server-ip>/` from another system on the same network. Make sure the firewall allows inbound TCP traffic on port `80`.

## Windows service

Run from an elevated PowerShell window:

```powershell
.\scripts\install_windows_service.ps1 -Port 80
```

To remove it:

```powershell
.\scripts\uninstall_windows_service.ps1
```

## Linux service

Run with sudo:

```bash
sudo PORT=80 ./scripts/install_linux_service.sh
```

To remove it:

```bash
sudo ./scripts/uninstall_linux_service.sh
```

## API

Health check:

```bash
curl http://<server-ip>/health
```

Analyze with multipart upload:

```bash
curl -X POST http://<server-ip>/api/analyze \
  -F host=localhost \
  -F port=11434 \
  -F model=llama3.1 \
  -F jd_file=@job-description.pdf \
  -F resumes=@resume-1.pdf \
  -F resumes=@resume-2.docx
```
