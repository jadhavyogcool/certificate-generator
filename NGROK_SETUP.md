# Setting Up Ngrok for External Access

## What is Ngrok?

Ngrok creates a secure tunnel to expose your local Flask app to the internet with a public URL. This allows anyone to access your certificate generator from anywhere!

---

## Step 1: Install Ngrok

### Option A: Download Directly (Recommended)

1. **Visit:** https://ngrok.com/download
2. **Download** the Windows version (ZIP file)
3. **Extract** the `ngrok.exe` file to a folder (e.g., `C:\ngrok\`)
4. **Add to PATH** (optional but recommended):
   - Right-click "This PC" → Properties → Advanced System Settings
   - Click "Environment Variables"
   - Under "System variables", find "Path" and click Edit
   - Click "New" and add `C:\ngrok\` (or wherever you extracted it)
   - Click OK on all dialogs

### Option B: Using Chocolatey (If you have it)

```powershell
choco install ngrok
```

---

## Step 2: Sign Up for Ngrok (Free)

1. **Create account:** https://dashboard.ngrok.com/signup
2. **Get your authtoken:** https://dashboard.ngrok.com/get-started/your-authtoken
3. **Copy the authtoken** (looks like: `2abc123def456ghi789jkl`)

---

## Step 3: Authenticate Ngrok

Open a **new PowerShell window** and run:

```powershell
ngrok config add-authtoken YOUR_AUTHTOKEN_HERE
```

Replace `YOUR_AUTHTOKEN_HERE` with your actual token.

---

## Step 4: Run Your Flask App

**Keep your Flask app running** in the current terminal:

```powershell
.\venv\Scripts\activate.ps1
python app.py
```

Your app should be running on `http://localhost:5000`

---

## Step 5: Start Ngrok Tunnel

Open a **NEW PowerShell window** and run:

```powershell
ngrok http 5000
```

You should see output like this:

```
Session Status                online
Account                       Your Name (Plan: Free)
Version                       3.x.x
Region                        United States (us)
Latency                       -
Web Interface                 http://127.0.0.1:4040
Forwarding                    https://abc123.ngrok-free.app -> http://localhost:5000

Connections                   ttl     opn     rt1     rt5     p50     p90
                              0       0       0.00    0.00    0.00    0.00
```

---

## Step 6: Share Your App! 🎉

Look for the **Forwarding** line. You'll see a URL like:

```
https://abc123.ngrok-free.app
```

**This is your public URL!** Share it with anyone to let them use your certificate generator from anywhere in the world.

---

## Important Notes

### ✅ **Advantages:**
- Works immediately - no complex setup
- Free tier available
- HTTPS by default (secure)
- Works from anywhere

### ⚠️ **Limitations (Free Tier):**
- URL changes every time you restart ngrok
- Session expires after 2 hours (just restart ngrok)
- Limited bandwidth
- Shows ngrok warning page before accessing your app

### 🔒 **Security Tips:**
- Only share the URL with trusted users
- Stop ngrok when not in use (Ctrl+C)
- Don't commit your authtoken to Git

---

## Keeping Both Running

You need **TWO terminals**:

**Terminal 1:** Flask app
```powershell
cd c:\Users\Dell\Certificate-Generator\certificate-generator
.\venv\Scripts\activate.ps1
python app.py
```

**Terminal 2:** Ngrok tunnel
```powershell
ngrok http 5000
```

---

## Stopping Everything

**To stop ngrok:**
- Press `Ctrl+C` in the ngrok terminal

**To stop Flask:**
- Press `Ctrl+C` in the Flask terminal

---

## Troubleshooting

### "ngrok not recognized"
- Make sure you extracted `ngrok.exe` and added it to PATH
- Or run it from the folder where you extracted it: `.\ngrok.exe http 5000`

### "Failed to authenticate"
- Run: `ngrok config add-authtoken YOUR_TOKEN`
- Make sure you copied the full token from ngrok dashboard

### Ngrok says "tunnel not found"
- Make sure Flask is running on port 5000 first
- Then start ngrok in a separate terminal

---

## Upgrading to Ngrok Paid (Optional)

For production use, consider ngrok's paid plans which offer:
- Custom domains (static URLs)
- No session limits
- More bandwidth
- Remove ngrok branding

Visit: https://ngrok.com/pricing

---

## Quick Reference

```powershell
# Authenticate (one time only)
ngrok config add-authtoken YOUR_TOKEN

# Start tunnel
ngrok http 5000

# Check status
# Visit http://localhost:4040 in browser
```

That's it! You're ready to share your certificate generator with the world! 🌍✨
