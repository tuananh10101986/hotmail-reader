# Hotmail Reader

Fullstack app to read Hotmail/Outlook inbox using OAuth2 + Microsoft Graph API.

## Run locally
1. Copy `backend/.env.example` -> `backend/.env` and fill in credentials from Azure app registration.
2. Run backend:
   ```bash
   cd backend
   npm install
   node server.js
   ```
3. Run frontend:
   ```bash
   cd frontend
   npm install
   npm run dev
   ```

## Run with Docker
```bash
docker-compose up --build
```
Frontend: http://localhost:5173  
Backend: http://localhost:3000
