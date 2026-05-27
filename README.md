# Slide Study

Browser-first PDF study workspace built with React, Vite, PDF.js, and Google Drive Picker.

## Local Development

Create `webApp/.env.local`:

```env
VITE_GOOGLE_CLIENT_ID=your-google-oauth-client-id.apps.googleusercontent.com
VITE_GOOGLE_APP_ID=your-google-cloud-project-number
VITE_GOOGLE_API_KEY=your-google-api-key
```

Run locally:

```powershell
npm install
npm run dev -- --host 127.0.0.1 --port 5179
```

## GitHub Pages

This repo deploys to:

```text
https://bebrave4321.github.io/slide-study/
```

Repository settings required for Actions:

- Variable `VITE_GOOGLE_CLIENT_ID`
- Variable `VITE_GOOGLE_APP_ID`
- Secret `VITE_GOOGLE_API_KEY`

Google Cloud Console must allow:

- OAuth JavaScript origin: `https://bebrave4321.github.io`
- API key website restriction: `https://bebrave4321.github.io/slide-study/*`
