# AuditEase

AuditEase is a secure full-stack audit document management web application.

---

## 🚀 How to Install and Run

### 1. Prerequisites

This project uses native Node modules (`bcrypt`, `better-sqlite3`) that require 
compilation tools on first install. Also, Node.js and SQLite are required.

* **Arch Linux**:
  ```bash
  sudo pacman -Syu --needed base-devel python nodejs npm sqlite
  ```
* **Debian / Ubuntu**:
  ```bash
  sudo apt update && sudo apt install -y curl build-essential python3 sqlite3
  curl -fsSL https://deb.nodesource.com/setup_18.x | sudo -E bash -
  sudo apt install -y nodejs
  ```
* **macOS**:
  ```bash
  xcode-select --install
  brew install node
  ```
* **Windows**:
  Download and run the official [Node.js MSI installer](https://nodejs.org/). Make sure to check the box that automatically installs build tools for native modules.
  Alternatively, install "Desktop development with C++" via Visual Studio Build Tools, or run: `npm install --global windows-build-tools` (run as Administrator).

---

### 2. Setup and Run the App

1. **Installation**:
   ```bash
   npm install
   ```
   If prompted with "packages have install scripts not yet covered by allowScripts", this is npm's security feature for native modules. Run:
   ```bash
   npm approve-scripts bcrypt better-sqlite3
   npm install
   ```
   This is expected on every fresh machine the first time — it's not an error.

   **Troubleshooting Note**: If `npm install` finishes but fails to build `bcrypt` or `better-sqlite3`, ensure you have the correct system prerequisites installed (see above).

   **Installing on a New Machine**
   Prefer this over `npm install` for consistent results across machines:
   ```bash
   npm ci
   ```
   If you get a "package-lock.json is out of sync" error, run `npm install` once to regenerate the lockfile, commit the updated lockfile, then others should use `npm ci` going forward.
2. **Configure environment variables**:
   Create a `.env` file in the root directory:
   ```bash
   cp .env.example .env
   ```
   Generate a cryptographically secure 64-character encryption key and place it under `ENCRYPTION_KEY` in `.env`:
   ```bash
   node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"
   ```
3. **Start the server**:
   ```bash
   npm start
   ```
   The application will run at **`http://localhost:3000`**.

---

## 👥 User Management

AuditEase includes CLI scripts to manage accounts directly from the terminal.

### 1. List All Users
Display a directory of all registered users:
```bash
npm run list-users
```

### 2. Add a New User
Create a new user with a hashed password:
```bash
npm run add-user -- --name "John Doe" --username johndoe --password securepass123 --role auditor
```
types of roles: company(default) and auditor

### 3. Change a User's Password
Update the password of an existing user securely:
```bash
npm run change-password -- --username johndoe --password newsecurepass123
```

### 4. Update a User's Role
Change the role of a user (valid roles are `company` and `auditor`):
```bash
npm run update-role -- --username johndoe --role auditor
```

### 5. Delete a User
Remove a user from the system safely (fails if they have active document records linked to them):
```bash
npm run delete-user -- --username johndoe
```

---

## 🔒 Gitignore & Security
The following directory and local database files are ignored in [.gitignore](file:///home/ash/Projects/AuditEase/.gitignore) so they do not get shipped to production:
* `node_modules/` (dependencies)
* `.env` (secrets)
* `backend/storage/vault/*` (encrypted files)
* `auditease.db*` (SQLite databases and transaction logs)
