Knowledge folder - drop your reference documents here.

Drop any of these file types into this folder:

  - .txt or .md  (read directly)
  - .docx        (text extracted via mammoth)

What goes here:
  - CoVE Process Writing Golden Rules (.docx)
  - RMIT Tips for Process Editors and Champions (.docx)
  - Nintex Process Writing Techniques (.docx)
  - Any other reference material you want the AI to be able to cite

The server reads everything in this folder on startup and includes the text
in the system prompt sent to OpenAI on every chat request. The AI can then
quote and cite specific documents by name in its responses.

After you add or change a file, restart the server (Ctrl+C, then `npm run dev`)
so the new content gets loaded.

The /api/health endpoint shows how many characters of knowledge are currently
loaded. Visit http://localhost:3001/api/health to confirm your files were
picked up.

This README itself won't be loaded - the server skips dotfiles and ignores
this filename because it's `.txt` but explicitly named README. (Actually it
WILL be loaded since the loader doesn't have name-based exclusions. If you
want it ignored, rename to README.md or move out of this folder.)

NOTE: knowledge files are NOT committed to git (see .gitignore).
