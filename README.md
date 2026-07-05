# Relay Local · LAN Share Hub

**One Python file. Zero dependencies. Share files and chat across every device on your WiFi.**

A self-hosted transfer station with a retro terminal aesthetic — file browser, chat-style messaging, previews, search, and batch operations. No WeChat file assistant, no QQ, no app install. Open a URL and go.

![Python 3](https://img.shields.io/badge/python-3.8+-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Dependencies](https://img.shields.io/badge/deps-none-brightgreen)

<p align="center">
  <img src="assets/hub-icon.svg" width="72" alt="Relay Local icon">
</p>

## Why this exists

Phones and PCs live on the same WiFi, but moving files between them is still painful: WeChat compresses images, file transfer assistants cap size, and every tool wants you inside their app ecosystem.

Relay Local runs on your PC. Any device on the LAN opens one page — send photos, videos, zips, text, search history, browse your home folder. **Trusted LAN only** (no auth by default). Keep the port on your local network.

## Quick start

```bash
git clone https://github.com/metahubaifeel/lan-share.git
cd lan-share
python3 server.py
```

Or:

```bash
chmod +x start.sh
./start.sh
```

Open `http://<your-pc-ip>:8888` on your phone (same WiFi). The top bar shows the URL — tap to copy.

## Features

| Area | What you get |
|------|----------------|
| **Chat** | WeChat-style thread: text, images, video, audio, zip cards, inline preview |
| **Multi-attach** | Select many files at once; queue bar with thumbnails; pack as zip or send separately |
| **Upload progress** | Full-width banner with %, speed, and clear error messages |
| **Download** | Progress for small files; large files delegate to the browser download manager |
| **Files** | Browse `~/`, filter by type, sort, breadcrumb navigation, desktop side preview |
| **Phone gallery** | Year-grouped horizontal strips for photos/videos from the inbox |
| **Search** | `Ctrl+K` — chat history + filename search across the serve tree |
| **Multi-select** | Long-press (mobile) or toggle mode — batch download, delete, preview strip |
| **Desktop extras** | Drag-and-drop upload, paste images from clipboard, desktop notifications |
| **Previews** | Images, video, audio, Markdown, PDF, text, Word (.docx), HEIC conversion |
| **Polish** | Terminal gold-on-black UI, tap feedback, optional 8-bit SFX, reduced-motion safe |

## Configuration

Environment variables (all optional):

| Variable | Default | Description |
|----------|---------|-------------|
| `HUB_PORT` | `8888` | Listen port |
| `HUB_BIND` | `0.0.0.0` | Bind address |
| `HUB_SERVE_DIR` | `~/` | Root for file browsing & downloads |
| `HUB_UPLOAD_DIR` | `~/lan-share/inbox/` | Where phone uploads land |
| `HUB_DB_PATH` | `~/.config/lan-share/messages.db` | Chat SQLite database |
| `HUB_MAX_UPLOAD_MB` | `0` (unlimited) | Max upload size in MB |
| `HUB_PUBLIC_URL` | *(empty)* | If set, enables cloud branding UI (advanced/self-hosted) |

CLI:

```bash
python3 server.py [port] [serve_dir]
```

## Auto-start (Linux, systemd user)

```bash
mkdir -p ~/lan-share
cp -r . ~/lan-share/   # or clone into ~/lan-share

mkdir -p ~/.config/systemd/user
cp examples/systemd/lan-share.service ~/.config/systemd/user/
# Edit paths in the unit file if needed

systemctl --user daemon-reload
systemctl --user enable --now lan-share.service
```

## Security note

Default mode exposes your serve directory to anyone on the LAN who knows the URL. Use only on networks you trust. Do **not** port-forward to the public internet without adding authentication (reverse proxy + HTTPS + Basic Auth, etc.).

Path traversal is blocked; dot-directories like `.git` are skipped during browse/search.

## Project layout

```
lan-share/
├── server.py              # HTTP server + SQLite + embedded SPA
├── start.sh               # Local launcher
├── assets/                # Icons
├── examples/
│   └── systemd/           # User service template
├── LICENSE
└── README.md
```

## API (short)

| Method | Path | Description |
|--------|------|-------------|
| GET | `/api/info` | Hub URL, mode, upload dir |
| GET | `/api/browse?path=` | Directory listing |
| GET | `/api/received` | Inbox listing |
| GET | `/api/download/<path>` | File download (Range / 206) |
| POST | `/api/upload` | Multipart upload |
| GET | `/api/messages` | Chat messages (poll with `?since=`) |
| POST | `/api/messages` | Send message |
| GET | `/api/search?q=` | Search chat + files |

## License

MIT — see [LICENSE](LICENSE).

---

Made for people who want **one clean transfer station on their own network**, not another app ecosystem.
