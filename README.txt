AI Internship Candidate Dashboard — beginner setup

WHAT THIS STARTER DOES
- Reads the 100 candidate rows from Internship_Candidates_100.xlsx.
- Displays candidates in a local dashboard.
- Supports search and status filtering.
- Opens valid GitHub links.
- Lets you update interview date and selection status and saves changes to Excel.
- Includes an "Explain with AI" demo placeholder. It is NOT connected to a real AI model yet.

SETUP ON WINDOWS
1. Install Python 3 from https://www.python.org/downloads/ if it is not installed. During installation, tick "Add Python to PATH".
2. Install Visual Studio Code from https://code.visualstudio.com/ if needed.
3. Extract this ZIP file to your Desktop.
4. Copy Internship_Candidates_100.xlsx into the extracted folder, next to app.py.
5. Open VS Code. Choose File > Open Folder and select the extracted internship_candidate_system folder.
6. Choose Terminal > New Terminal.
7. Run: py -m pip install -r requirements.txt
8. Run: py app.py
9. In the terminal, find the local address, usually http://127.0.0.1:5000
10. Hold Ctrl and click the address, or copy it into Chrome.
11. Keep the terminal open while using the dashboard. Press Ctrl+C in the terminal to stop it.

EXCEL COLUMN ORDER EXPECTED
A Candidate ID
B Candidate Name
C Email
D Project Title
E GitHub Link
F Applied Date
G Interview Date
H AI Project Explanation
I Selection Status

IMPORTANT
- Use only real candidate data you are authorized to handle.
- The supplied names/emails are placeholders.
- The app is a local demo, not a secure public production service.
- The AI button does not analyse code yet. To add real AI, connect an AI provider on the Python backend and safely fetch/read accessible repository README/source files. Never put an API key in HTML or JavaScript.
- The app expects dates in the Excel Interview Date column to be compatible with the browser date input; after saving, dates are stored in ISO YYYY-MM-DD format.
