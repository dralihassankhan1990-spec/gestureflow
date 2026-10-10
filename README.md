# GestureFlow — Hand-Gesture PowerPoint Control

Control PowerPoint with hand gestures through your webcam. No special hardware — MediaPipe hand tracking in the browser sends gesture commands over a local WebSocket to a Node bridge that drives PowerPoint on Windows.

## Gestures

| Gesture | Action | How |
|---|---|---|
| ✋ Open palm | Start slideshow | Hold 1 second |
| ✊ Closed fist | Close presentation | Hold 0.8 seconds |
| ☝️ Pointing | Laser pointer toggle | Hold 0.6 seconds, then move hand to point |
| ➡️ Swipe right | Next slide | Quick horizontal swipe |
| ⬅️ Swipe left | Previous slide | Quick horizontal swipe |
| 🖐️ Spread (thumb + index) | Zoom in | Spread fingers |
| 🤏 Pinch (thumb + index) | Zoom out | Pinch fingers |

Hold-times are configurable in `App.tsx` (`holdRequired`). The app also ships two interactive demo modes — **Particles** and **Carousel** — that respond to the same gestures without PowerPoint.

## How It Works

```
Webcam → MediaPipe hand tracking (browser) → WebSocket (ws://localhost:3001)
       → bridge.cjs (Node) → PowerShell COM handler (ps_handler.ps1) → PowerPoint
```

- `bridge.cjs` — local WebSocket server; routes gesture events to PowerPoint. Uses `robotjs` for low-latency cursor input when available, otherwise falls back to the persistent PowerShell handler.
- `ps_handler.ps1` — persistent PowerShell process talking to PowerPoint via COM.
- `powerpoint-addin/` — companion Office add-in (separate sub-project with its own `package.json`).

## Tech Stack

- React 19 + TypeScript + Vite
- `@mediapipe/tasks-vision` (hand landmark detection)
- `ws` (WebSocket server), `robotjs` (optional native input control), `concurrently`
- PowerShell + COM automation for PowerPoint

## Getting Started

**Prerequisites:** Node.js 18+, Windows with PowerPoint installed (for presentation control; demo modes work anywhere).

```bash
# Install dependencies
npm install

# Run the app + bridge together (http://localhost:3000)
npm run dev

# Production build (frontend only)
npm run build
```

> **Windows only:** PowerPoint control requires the PowerShell COM bridge (`bridge.cjs` + `ps_handler.ps1`), which `npm run dev` starts automatically alongside Vite. `robotjs` is an optional dependency — if its native build fails on your machine, the bridge falls back to PowerShell-only control.

For gesture tips, see [GESTURES_GUIDE.md](GESTURES_GUIDE.md).

## Author

Muhammad Bilal
