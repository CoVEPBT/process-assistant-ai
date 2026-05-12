# Process Assistant AI

A version of the CoVE Process Writing Assistant with **inline AI** (OpenAI gpt-4o-mini) instead of the external Copilot agent. Lets users get conversational help and AI-driven enhancement without leaving the tool.

This is a **separate version** from the SharePoint-deployed v2 (`workflow-generator/rmit-process-writing-assistant-v2.html`) — that one is unchanged and continues to use the M365 Copilot agent for staff use.

## Architecture

Mirrors comms-engine: Express server holds the OpenAI API key in `.env`, frontend calls the server, server calls OpenAI. Keys never reach the browser.

```
process-assistant-ai/
├── server/
│   └── index.js        Express server with /api/chat endpoint
├── public/
│   └── index.html      Frontend (copy of workflow-generator v2 with AI)
├── package.json
├── .env.example
└── README.md
```

## Getting started

```bash
# 1. Install dependencies
cd process-assistant-ai
npm install

# 2. Set up .env
cp .env.example .env
# Edit .env and paste your OpenAI API key into OPENAI_API_KEY

# 3. Run the server
npm run dev
# Open http://localhost:3001
```

## How it works

- The frontend (HTML) handles all the rule-based analysis client-side: parsing, AU spelling, auto-fixes, rules engine, exports. Same as v2.
- When the user clicks "AI Review" or chats with the AI panel, the frontend POSTs to `/api/chat` on the local server.
- The server calls OpenAI's chat completions API with the persona prompt + the user's document + their question.
- The response comes back to the frontend and renders as a chat reply.

## Cost

`gpt-4o-mini` is inexpensive: ~$0.15 per 1M input tokens, $0.60 per 1M output tokens. A typical process review (2K input + 1K output) costs roughly $0.001 per call. Pilot usage of 100 calls/day = ~$0.10/day.

## Security

- `.env` is gitignored and must never be committed
- The API key only exists on the machine running the server — not in the browser, not in any deployed file
- For testing on your own machine, run `npm run dev` and access http://localhost:3001
- For wider rollout, the server needs to be hosted somewhere accessible to your users (similar hosting question as comms-engine + VAL — the same RMIT IT conversation applies)


## Knowledge files

Drop reference documents (CoVE Golden Rules, RMIT Tips, Nintex Techniques, anything else) into the `knowledge/` folder. Supported types: `.txt`, `.md`, `.docx`. The server reads them on startup and includes the content in the system prompt so the AI can cite them.

```bash
# Example layout
knowledge/
  CoVE-Golden-Rules.docx
  RMIT-Tips-for-Process-Editors.docx
  Nintex-Process-Writing-Techniques.docx
```

After adding or changing files, restart the server (`Ctrl+C` then `npm run dev`). The startup log shows `[knowledge] loaded ...` for each file picked up. Visit `http://localhost:3001/api/health` to see the total character count of loaded knowledge.

Knowledge files are gitignored — keep RMIT-internal content out of any public repo.

## Comparison with v2 (Copilot version)

|                       | v2 (Copilot)                    | v3 (AI)                          |
| --------------------- | ------------------------------- | -------------------------------- |
| AI provider           | M365 Copilot agent              | OpenAI gpt-4o-mini               |
| Where AI runs         | Microsoft tenant boundary       | Your OpenAI account              |
| Hosting requirement   | None — pure static HTML         | Node.js server (this repo)       |
| API key handling      | None — uses your M365 sign-in   | Server-side .env                 |
| Deployment            | GitHub Pages → SharePoint embed | Localhost or hosted Node server  |
| User experience       | Opens agent in new tab          | Inline chat, no tab switch       |
| Cost per use          | Bundled in M365 Copilot licence | Per-token OpenAI charges         |
| Suited for            | RMIT staff at scale             | Solo testing, small pilot group  |
